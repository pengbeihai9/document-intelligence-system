#!/usr/bin/env bash
set -euo pipefail

PROJECT_DIR="${PROJECT_DIR:-/root/doc_summary_lora}"
MODEL_NAME="${MODEL_NAME:-Qwen/Qwen2.5-7B-Instruct}"
DATA_FILE="${DATA_FILE:-$PROJECT_DIR/data/paper_summary_sft.jsonl}"
OUT_DIR="${OUT_DIR:-$PROJECT_DIR/outputs/qwen2_5_7b_paper_summary_lora}"

mkdir -p "$PROJECT_DIR" "$PROJECT_DIR/data" "$PROJECT_DIR/models" "$OUT_DIR"
cd "$PROJECT_DIR"

export HF_HOME="${HF_HOME:-$PROJECT_DIR/models/hf_home}"
export TRANSFORMERS_CACHE="${TRANSFORMERS_CACHE:-$PROJECT_DIR/models/transformers}"

if [ ! -d ".venv" ]; then
  python3 -m venv .venv
fi
source .venv/bin/activate

python -m pip install --upgrade pip
python -m pip install "torch" "transformers>=4.43" "datasets" "accelerate" "peft" "trl" "bitsandbytes" "sentencepiece"

cat > "$PROJECT_DIR/train_lora.py" <<'PY'
import os

from datasets import load_dataset
from peft import LoraConfig
from transformers import AutoModelForCausalLM, AutoTokenizer, BitsAndBytesConfig
from trl import SFTConfig, SFTTrainer


model_name = os.environ.get("MODEL_NAME", "Qwen/Qwen2.5-7B-Instruct")
data_file = os.environ["DATA_FILE"]
out_dir = os.environ["OUT_DIR"]

tokenizer = AutoTokenizer.from_pretrained(model_name, trust_remote_code=True)
if tokenizer.pad_token is None:
    tokenizer.pad_token = tokenizer.eos_token


def format_chat(example):
    return tokenizer.apply_chat_template(example["messages"], tokenize=False, add_generation_prompt=False)


dataset = load_dataset("json", data_files=data_file, split="train")
dataset = dataset.map(lambda item: {"text": format_chat(item)}, remove_columns=dataset.column_names)

quant_config = BitsAndBytesConfig(
    load_in_4bit=True,
    bnb_4bit_quant_type="nf4",
    bnb_4bit_compute_dtype="bfloat16",
    bnb_4bit_use_double_quant=True,
)

model = AutoModelForCausalLM.from_pretrained(
    model_name,
    quantization_config=quant_config,
    device_map="auto",
    trust_remote_code=True,
)
model.config.use_cache = False

peft_config = LoraConfig(
    r=16,
    lora_alpha=32,
    lora_dropout=0.05,
    bias="none",
    task_type="CAUSAL_LM",
    target_modules=["q_proj", "k_proj", "v_proj", "o_proj", "gate_proj", "up_proj", "down_proj"],
)

training_args = SFTConfig(
    output_dir=out_dir,
    num_train_epochs=float(os.environ.get("EPOCHS", "2")),
    per_device_train_batch_size=int(os.environ.get("BATCH_SIZE", "1")),
    gradient_accumulation_steps=int(os.environ.get("GRAD_ACCUM", "8")),
    learning_rate=float(os.environ.get("LR", "2e-4")),
    logging_steps=5,
    save_steps=50,
    save_total_limit=2,
    bf16=True,
    max_length=int(os.environ.get("MAX_SEQ_LENGTH", "4096")),
    packing=False,
    report_to="none",
)

trainer = SFTTrainer(
    model=model,
    processing_class=tokenizer,
    train_dataset=dataset,
    peft_config=peft_config,
    args=training_args,
)
trainer.train()
trainer.save_model(out_dir)
tokenizer.save_pretrained(out_dir)
print(f"LoRA adapter saved to {out_dir}")
PY

MODEL_NAME="$MODEL_NAME" DATA_FILE="$DATA_FILE" OUT_DIR="$OUT_DIR" python "$PROJECT_DIR/train_lora.py"
