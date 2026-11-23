import json
import re
import sys
from contextlib import closing, contextmanager
from dataclasses import dataclass
from pathlib import Path


RESOURCE_DIR = Path(__file__).resolve().parents[1]
APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else RESOURCE_DIR
MODELS_DIR = APP_DIR / "models"
PROMPT_PATH = APP_DIR / "system_prompt.txt"

STATUS_LABELS = {
    "in_customs": "In customs",
    "shipped": "Shipped",
    "ready_to_ship": "Ready to ship",
    "unknown": "Unknown",
}
CLASSIFICATION_SCHEMA = {
    "type": "object",
    "properties": {"status": {"type": "string", "enum": list(STATUS_LABELS)}},
    "required": ["status"],
    "additionalProperties": False,
}
OUTPUT_FORMAT = "\n\nFor this classification, return a JSON object containing only the status field."


@dataclass(frozen=True)
class Classification:
    status: str


def load_system_prompt():
    source = PROMPT_PATH
    if not source.exists():
        source = RESOURCE_DIR / PROMPT_PATH.name
        if not source.exists():
            return ""

    text = source.read_text(encoding="utf-8-sig").strip()
    # Accept the existing triple-quoted assignment as well as plain text.
    assignment = re.fullmatch(r"system_prompt\s*=\s*('''|\"\"\")(.*?)\1", text, re.DOTALL)
    return assignment[2].strip() if assignment else text


def order_conversations(order):
    if "conversations" in order:
        return order["conversations"]
    return [order["emails"]] if order["emails"] else []


@contextmanager
def load_model(path, is_cancelled):
    if not path.is_file():
        raise ValueError(f"Move a GGUF model into {MODELS_DIR} and select it again.")
    try:
        from llama_cpp import Llama, llama_cpp
    except (ImportError, OSError, RuntimeError) as error:
        if getattr(sys, "frozen", False):
            hint = "Rebuild the app with llama-cpp-python installed."
        else:
            hint = "Install llama-cpp-python in the app's Python environment."
        raise RuntimeError(f"The local model runtime could not be loaded. {hint}\n\n{error}") from error
    # The bindings' quiet mode redirects console descriptors, which a windowed build can lack.
    verbose = sys.stdout is None or sys.stderr is None
    with closing(Llama(model_path=str(path), n_ctx=16384, n_gpu_layers=0, verbose=verbose)) as model:
        # Keep the callback alive until the native context is closed.
        abort_callback = llama_cpp.ggml_abort_callback(lambda _: is_cancelled())
        llama_cpp.llama_set_abort_callback(model.ctx, abort_callback, None)
        yield model


def classify_order(model, prompt, order):
    conversation_ids = {
        (email["store_id"], email["entry_id"]): f"c{index}"
        for index, conversation in enumerate(order_conversations(order), 1)
        for email in conversation
    }
    messages = [
        {
            "conversation_id": conversation_ids.get((email["store_id"], email["entry_id"])),
            "sender": email["sender"],
            "received_at": email["received_at"],
            "subject": email["subject"],
            "body": email["body"],
        }
        for email in order["emails"]
    ]
    line = order.get("LINE")
    payload = {
        "target": {"po": str(order["PO"]).strip(), "line": None if line is None else str(line).strip() or None},
        "messages": messages,
    }
    try:
        answer = model.create_chat_completion(
            messages=[
                {"role": "system", "content": prompt + OUTPUT_FORMAT},
                {"role": "user", "content": json.dumps(payload, ensure_ascii=False)},
            ],
            response_format={"type": "json_object", "schema": CLASSIFICATION_SCHEMA},
            temperature=0,
            max_tokens=512,
        )
    except ValueError as error:
        if "context window" in str(error):
            raise ValueError("The emails for this PO/line exceed the model's context window.") from error
        raise

    choice = answer["choices"][0]
    if choice.get("finish_reason") == "length":
        raise ValueError("The model ran out of space before finishing its answer.")
    try:
        result = json.loads(choice["message"]["content"])
    except (TypeError, json.JSONDecodeError) as error:
        raise ValueError("The model returned invalid classification JSON.") from error

    if not isinstance(result, dict) or set(result) != {"status"}:
        raise ValueError("The model response must contain a single order status.")
    status = result["status"]
    if not isinstance(status, str) or status not in STATUS_LABELS:
        raise ValueError("The model returned an unsupported order status.")
    return Classification(status)
