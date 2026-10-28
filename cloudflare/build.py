"""Cloudflare 배포용 번들 생성: python cloudflare/build.py

build/src/       worker.py + medical_platform 패키지 (+ 템플릿·아이콘을 묶은 _embedded.py)
build/public/    정적 파일 (/static/... 로 Cloudflare Static Assets 가 직접 서빙)
"""
import pathlib
import shutil

HERE = pathlib.Path(__file__).resolve().parent
PKG = HERE.parent / "medical_platform"
OUT = HERE / "build"


def main():
    shutil.rmtree(OUT, ignore_errors=True)
    src = OUT / "src" / "medical_platform"
    src.mkdir(parents=True)
    for py in PKG.glob("*.py"):
        shutil.copy2(py, src / py.name)
    shutil.copy2(HERE / "worker.py", OUT / "src" / "worker.py")

    templates = {p.relative_to(PKG / "templates").as_posix(): p.read_text(encoding="utf-8")
                 for p in sorted((PKG / "templates").rglob("*.html"))}
    icons = {p.stem: p.read_text(encoding="utf-8")
             for p in sorted((PKG / "static" / "vendor" / "lucide" / "icons").glob("*.svg"))}
    (src / "_embedded.py").write_text(
        '"""자동 생성 파일 (cloudflare/build.py) — 직접 수정하지 마세요."""\n'
        f"TEMPLATES = {templates!r}\nICONS = {icons!r}\n", encoding="utf-8")

    shutil.copytree(PKG / "static", OUT / "public" / "static",
                    ignore=shutil.ignore_patterns("icons"))
    print(f"built {len(templates)} templates, {len(icons)} icons → {OUT}")


if __name__ == "__main__":
    main()
