from __future__ import annotations

import argparse
import os
import shlex
import shutil
import subprocess
import tempfile
from pathlib import Path

try:
    from .textconv import canonical_diff
except ImportError:
    from textconv import canonical_diff


def find_drift(docx_path, rebuild_cmd: list[str]) -> list[str]:
    """Bandingkan dokumen kerja dengan hasil rebuild manual yang diminta.

    Dokumen kerja disalin ke baseline sementara sebelum `rebuild_cmd` dijalankan.
    Perintah rebuild dijalankan di direktori dokumen dengan variabel lingkungan
    ``DOCX_ENGINE_TARGET`` yang menunjuk ke dokumen kerja; baseline tidak pernah
    ditulis kembali ke dokumen tersebut.
    """
    target = Path(docx_path).resolve()
    with tempfile.TemporaryDirectory(prefix="docx-drift-") as unpacked:
        baseline = Path(unpacked) / target.name
        shutil.copy2(target, baseline)

        env = os.environ.copy()
        env["DOCX_ENGINE_TARGET"] = str(target)
        proc = subprocess.run(
            list(rebuild_cmd),
            cwd=str(target.parent),
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        if proc.returncode != 0:
            stderr = proc.stderr
            if isinstance(stderr, bytes):
                stderr = stderr.decode("utf-8", errors="replace")
            raise RuntimeError(
                f"perintah rebuild gagal dengan kode {proc.returncode}: {stderr}"
            )

        return canonical_diff(baseline, target)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Deteksi drift kanonik setelah rebuild manual dokumen DOCX",
    )
    parser.add_argument("docx", help="Path dokumen DOCX yang diperiksa")
    parser.add_argument(
        "--rebuild",
        required=True,
        help="Perintah rebuild (dipecah dengan shlex.split)",
    )
    args = parser.parse_args()

    findings = find_drift(args.docx, shlex.split(args.rebuild))
    for finding in findings:
        print(finding)
    raise SystemExit(1 if findings else 0)


if __name__ == "__main__":
    main()
