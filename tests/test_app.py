import unittest
from unittest.mock import patch

import requests

from app import app
from app import build_neo_workbook
from app import neo_request_error
from app import xafiro_request_error


class AppRoutesTestCase(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.client = app.test_client()

    def test_index_is_available(self):
        response = self.client.get("/")

        self.assertEqual(response.status_code, 200)
        self.assertIn(b"Contabilidad Chancay", response.data)
        self.assertIn(b"carga.json", response.data)
        self.assertIn(b"check.json", response.data)

    def test_animation_files_are_available(self):
        loading = self.client.get("/animaciones/carga.json")
        success = self.client.get("/animaciones/check.json")

        self.assertEqual(loading.status_code, 200)
        self.assertEqual(success.status_code, 200)
        self.assertIn(b"Download-lottie", loading.data)
        loading.close()
        success.close()

    def test_rejects_inverted_date_range(self):
        response = self.client.post(
            "/api/xafiro/export",
            json={"fecha_inicio": "2026-08-10", "fecha_fin": "2026-08-01"},
        )

        self.assertEqual(response.status_code, 400)
        self.assertIn("posterior", response.get_json()["error"])

    def test_rejects_invalid_date_format(self):
        response = self.client.post(
            "/api/xafiro/export",
            json={"fecha_inicio": "01/08/2026", "fecha_fin": "2026-08-10"},
        )

        self.assertEqual(response.status_code, 400)
        self.assertIn("AAAA-MM-DD", response.get_json()["error"])

    def test_explains_windows_socket_block(self):
        error = xafiro_request_error(
            requests.ConnectionError("[WinError 10013] Intento de acceso a un socket no permitido"),
            "iniciar sesión",
        )

        self.assertIn("bloqueó", str(error))
        self.assertIn("firewall", str(error))

    def test_explains_neo_windows_socket_block(self):
        error = neo_request_error(
            requests.ConnectionError("[WinError 10013] Intento de acceso a un socket no permitido"),
            "iniciar sesión",
        )

        self.assertIn("Neo", str(error))
        self.assertIn("firewall", str(error))

    @patch("app.download_xafiro_report")
    def test_returns_report_as_attachment(self, download_report):
        download_report.return_value = (
            b"PK\x03\x04test-report",
            "facturacion_xafiro_2026-08-01_2026-08-02.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        response = self.client.post(
            "/api/xafiro/export",
            json={"fecha_inicio": "2026-08-01", "fecha_fin": "2026-08-02"},
        )

        self.assertEqual(response.status_code, 200)
        self.assertIn("attachment", response.headers["Content-Disposition"])
        self.assertTrue(response.data.startswith(b"PK\x03\x04"))

    @patch("app.download_neo_report")
    def test_returns_neo_report_as_attachment(self, download_report):
        download_report.return_value = (
            b"PK\x03\x04test-report",
            "facturacion_neo_2026-08-01_2026-08-02.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        response = self.client.post(
            "/api/neo/export",
            json={"fecha_inicio": "2026-08-01", "fecha_fin": "2026-08-02"},
        )

        self.assertEqual(response.status_code, 200)
        self.assertIn("facturacion_neo_2026-08-01_2026-08-02.xlsx", response.headers["Content-Disposition"])
        self.assertTrue(response.data.startswith(b"PK\x03\x04"))

    def test_builds_neo_workbook(self):
        content = build_neo_workbook(
            [
                {
                    "fecha_operacion": "2026-08-01",
                    "producto": "Ticket adulto",
                    "cantidad": 2,
                    "precio": 30,
                }
            ]
        )

        self.assertTrue(content.startswith(b"PK\x03\x04"))


if __name__ == "__main__":
    unittest.main()
