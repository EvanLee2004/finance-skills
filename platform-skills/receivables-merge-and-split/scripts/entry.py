from __future__ import annotations

import argparse
import json
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import Any

import yaml


def emit(message: str, progress: int) -> None:
    print(
        json.dumps(
            {"type": "progress", "message": message, "progress": progress, "state": "running"},
            ensure_ascii=False,
        ),
        flush=True,
    )


def load_json(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise TypeError(f"{path.name} 必须是 JSON 对象。")
    return value


def file_items(payload: dict[str, Any], role: str) -> list[dict[str, Any]]:
    value = payload.get("files", {}).get(role)
    if not value:
        return []
    items = value if isinstance(value, list) else [value]
    return [item for item in items if isinstance(item, dict) and item.get("local_path")]


def add_binding(command: list[str], binding: dict[str, Any], payload: dict[str, Any]) -> None:
    kind = binding["kind"]
    flag = binding.get("flag")
    if kind in {"file", "files", "input_dir"}:
        items = file_items(payload, binding["role"])
        if not items:
            return
        values = [str(Path(item["local_path"]).resolve()) for item in items]
        if kind == "input_dir":
            values = [str(Path(values[0]).parent)]
        if flag:
            command.append(flag)
        command.extend(values if kind == "files" else values[:1])
        return

    if kind == "parameter":
        value = payload.get("parameters", {}).get(binding["name"])
        if value in (None, "", []):
            return
        if binding.get("boolean"):
            if bool(value) and flag:
                command.append(flag)
            return
        if flag:
            command.append(flag)
        if isinstance(value, list):
            command.extend(str(item) for item in value)
        else:
            command.append(str(value))
        return

    if kind == "static":
        command.extend(str(item) for item in binding.get("values", []))
        return

    raise ValueError(f"不支持的桥接参数类型：{kind}")


def archive_directory(source: Path, target: Path) -> None:
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(source.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(source))


def collect_outputs(output_dir: Path, config: dict[str, Any]) -> list[dict[str, str]]:
    output = config["output"]
    target = output_dir / output["path"]
    if output["type"] == "file":
        candidates = [target] if target.is_file() else []
        for pattern in output.get("additional_globs", []):
            candidates.extend(output_dir.glob(pattern))
    else:
        candidates = [path for path in target.rglob("*") if path.is_file()]

    required_globs = output.get("required_globs", [])
    if required_globs:
        if not target.is_dir():
            raise RuntimeError(f"原 Skill 未生成必需的结果目录：{target.name}。")
        missing = [
            pattern
            for pattern in required_globs
            if not any(path.is_file() for path in target.glob(pattern))
        ]
        if missing:
            raise RuntimeError(
                "原 Skill 未生成必需的交付文件：" + "、".join(str(pattern) for pattern in missing)
            )

    if output.get("archive"):
        archive_path = output_dir / output.get("archive_name", "执行结果.zip")
        archive_directory(target, archive_path)
        candidates = [archive_path]

    unique: list[Path] = []
    seen: set[Path] = set()
    for candidate in candidates:
        resolved = candidate.resolve()
        if resolved.is_file() and resolved not in seen:
            unique.append(resolved)
            seen.add(resolved)
    if not unique:
        raise RuntimeError("原 Skill 执行成功，但没有在任务输出目录生成交付文件。")
    return [{"name": path.name, "path": str(path)} for path in unique]


def sanitized_failure(stderr: str, stdout: str) -> str:
    lines = [line.strip() for line in (stderr + "\n" + stdout).splitlines() if line.strip()]
    text = "；".join(lines[-8:])
    return text[:1200] or "原 Skill 返回了非零退出码。"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--request", required=True)
    parser.add_argument("--result", required=True)
    args = parser.parse_args()

    request_path = Path(args.request).resolve()
    result_path = Path(args.result).resolve()
    skill_dir = Path(__file__).resolve().parents[1]
    config = yaml.safe_load((skill_dir / "bridge.yaml").read_text(encoding="utf-8"))
    payload = load_json(request_path)
    output_dir = Path(payload["output_dir"]).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    entrypoint = (skill_dir / config["command"]).resolve()
    if not entrypoint.is_file() or not entrypoint.is_relative_to(skill_dir):
        raise RuntimeError("原 Skill 入口不存在或超出 Skill 包目录。")
    command = [sys.executable, str(entrypoint)]
    for binding in config.get("arguments", []):
        add_binding(command, binding, payload)

    output = config["output"]
    target = output_dir / output["path"]
    if output["type"] == "directory":
        target.mkdir(parents=True, exist_ok=True)
    else:
        target.parent.mkdir(parents=True, exist_ok=True)
    if output.get("flag"):
        command.extend([output["flag"], str(target)])

    emit("正在准备输入文件", 15)
    completed = subprocess.run(
        command,
        cwd=entrypoint.parent,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        check=False,
    )
    if completed.returncode:
        raise RuntimeError(sanitized_failure(completed.stderr, completed.stdout))

    emit("正在整理交付文件", 80)
    artifacts = collect_outputs(output_dir, config)
    result = {
        "status": "success",
        "summary": {
            "skill": config["name"],
            "output_count": len(artifacts),
            "message": config.get("success_message", "处理完成。"),
        },
        "output_files": artifacts,
        "warnings": [],
    }
    result_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    emit("任务处理完成", 95)


if __name__ == "__main__":
    main()
