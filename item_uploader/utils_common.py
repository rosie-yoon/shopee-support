# utils_common.py 내 환경변수 유틸 근처에 추가 (get_env, get_bool_env 다음 등)

from pathlib import Path
import re

def save_env_value(name: str, value: str, search_paths: Optional[list[Path]] = None) -> bool:
    """
    .env에 name=value를 저장(있으면 교체, 없으면 추가).
    - Cloud(읽기전용)에서는 실패할 수 있으므로 False 반환 가능.
    - 로컬 개발 편의용 유틸. 반환값으로 성공 여부만 알려줌.
    """
    name = str(name).strip()
    value = str(value)
    if not name:
        return False

    # 탐색 경로: 현재 파일 근처 → 상위 → CWD
    base = Path(__file__).resolve().parent
    candidates = search_paths or [base / ".env", base.parent / ".env", Path.cwd() / ".env"]

    env_path: Optional[Path] = None
    for p in candidates:
        try:
            if p.exists() and p.is_file():
                env_path = p
                break
        except Exception:
            continue
    if env_path is None:
        # 첫 후보에 새로 생성 시도
        env_path = candidates[0]

    try:
        lines = []
        if env_path.exists():
            lines = env_path.read_text(encoding="utf-8").splitlines()

        pattern = re.compile(rf"^\s*{re.escape(name)}\s*=\s*.*$")
        replaced = False
        for i, line in enumerate(lines):
            if pattern.match(line):
                lines[i] = f"{name}={value}"
                replaced = True
                break
        if not replaced:
            lines.append(f"{name}={value}")

        env_path.parent.mkdir(parents=True, exist_ok=True)
        env_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
        return True
    except Exception:
        # 쓰기 불가(예: 클라우드 읽기전용 등)
        return False
