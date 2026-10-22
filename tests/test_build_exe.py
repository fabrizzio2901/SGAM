import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import build_exe


class BuildExeTests(unittest.TestCase):
    def test_actualiza_alias_existente_despues_de_compilar(self):
        with tempfile.TemporaryDirectory() as temporal:
            raiz = Path(temporal)
            dist = raiz / "dist"
            dist.mkdir()
            (raiz / "SGAM.spec").write_text("", encoding="utf-8")
            (dist / "SGAM.exe").write_bytes(b"version nueva")
            alias = dist / "SGAM_actualizado.exe"
            alias.write_bytes(b"version anterior")
            with patch.object(build_exe, "BASE_DIR", raiz), \
                 patch.object(build_exe, "SPEC_FILE", raiz / "SGAM.spec"), \
                 patch.object(build_exe.platform, "system", return_value="Windows"), \
                 patch.object(build_exe, "_asegurar_pyinstaller"), \
                 patch.object(build_exe.subprocess, "run", return_value=SimpleNamespace(returncode=0)), \
                 patch.object(build_exe.sys, "argv", ["build_exe.py"]):
                build_exe.build()
            self.assertEqual(alias.read_bytes(), b"version nueva")


if __name__ == "__main__":
    unittest.main()
