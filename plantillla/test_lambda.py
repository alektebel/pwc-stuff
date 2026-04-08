"""
Local integration test for the Lambda + DocumentProcessor.

Mocks boto3 (DynamoDB + S3) so it runs without AWS credentials.
Writes the generated .docx to sample_output.docx for visual inspection.

Run:
    cd /home/diego/dev/pwc/plantillla
    python test_lambda.py
"""

import base64
import importlib
import json
import sys
import unittest
import zipfile
from io import BytesIO
from pathlib import Path
from unittest.mock import MagicMock, patch

# ── Path setup ────────────────────────────────────────────────────────────────
REPO_ROOT = Path(__file__).parent
LAMBDA_DIR = REPO_ROOT.parent / "lambda"
LAYER_DIR  = REPO_ROOT.parent / "layer" / "python"
TEMPLATE_PATH = REPO_ROOT / "plantilla-corrected - Copy.docx"
OUTPUT_PATH = REPO_ROOT / "sample_output.docx"

sys.path.insert(0, str(LAMBDA_DIR))
if LAYER_DIR.exists():
    sys.path.insert(0, str(LAYER_DIR))

# ── Sample content ────────────────────────────────────────────────────────────
SECTION_ALCANCE = """\
# 1. Alcance

Este informe cubre la **revisión de controles TI** para el ejercicio 2026.

## Objetivos

- Evaluar la eficacia de los controles de acceso lógico
- Revisar los procesos de gestión de cambios
- Verificar el cumplimiento de la política de seguridad

## Perímetro auditado

| Sistema | Entorno | Criticidad |
|---------|---------|------------|
| SAP ERP | Producción | Alta |
| Active Directory | Producción | Alta |
| Portal web | Pre-producción | Media |

> **Nota:** Se excluyen los sistemas de terceros sin acceso directo.
"""

SECTION_VALORACION = """\
# 2. Valoración

Se han identificado **3 hallazgos** durante el trabajo de campo.

## Hallazgo 1 — Gestión de accesos privilegiados

Los accesos de administrador no se revisan periódicamente.

### Riesgo

*Alto* — posible acceso no autorizado a datos sensibles.

### Recomendación

Implementar una revisión trimestral de accesos privilegiados con evidencia documentada.

## Hallazgo 2 — Gestión de parches

El 40 % de los servidores presentan parches pendientes de más de 90 días.

1. Inventariar todos los activos afectados
2. Priorizar según criticidad
3. Aplicar parches en ventana de mantenimiento mensual
"""

SECTION_CONCLUSIONES = """\
# 3. Conclusiones

La madurez del control interno en materia de TI es **media-baja**.

Se destaca positivamente:

- Existencia de un marco de políticas documentado
- Proceso de backup operativo con pruebas de restauración

Se requiere mejora urgente en:

- Gestión del ciclo de vida de accesos
- Programa de parcheo y vulnerabilidades

---

El equipo de auditoría agradece la colaboración prestada durante la revisión.
"""

SECTION_PROPUESTAS = """\
# 4. Propuestas de mejora

## Propuesta A — Plan de accesos

Desarrollar un procedimiento formal de revisión semestral de accesos que incluya:

- Responsable designado por sistema
- Evidencia de aprobación del negocio
- Registro de bajas inmediatas

## Propuesta B — Programa de parcheo

Establecer un *SLA de 30 días* para parches críticos y 90 días para el resto.

```
Criticidad CVSS ≥ 9.0  → parche en ≤ 7 días
Criticidad CVSS ≥ 7.0  → parche en ≤ 30 días
Resto                  → parche en ≤ 90 días
```

## Propuesta C — Cuadro de mando TI

Implantar un dashboard mensual con KPIs de seguridad para la Dirección.
"""

SECTIONS = {
    "1. Alcance":      SECTION_ALCANCE,
    "2. Valoracion":   SECTION_VALORACION,
    "3. Conclusiones": SECTION_CONCLUSIONES,
    "4. Propuestas":   SECTION_PROPUESTAS,
}

# ── Cover-page fields ─────────────────────────────────────────────────────────
COVER_FIELDS = {
    "audit_code":   "2601-0042",
    "audit_title":  "Auditoría de Controles TI",
    "uai":          "Tecnología",
    "date":         "09/04/2026",
    "recipients": [
        "D. Juan García López",
        "D. Pedro Martínez Ruiz",
        "Dª. Ana Fernández Soto",
    ],
    "audit_status": "BORRADOR",
}

# ── DynamoDB mock ─────────────────────────────────────────────────────────────
def _dynamo_get_item(Key, **_):
    sort_key = Key["report_sort"]
    content  = SECTIONS.get(sort_key, "")
    return {"Item": {"report_id": "test-001", "report_sort": sort_key,
                     "validated_content": content}}

# ── Tests ─────────────────────────────────────────────────────────────────────
class TestLambdaHandler(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        """Load template bytes once for all tests."""
        cls.template_bytes = TEMPLATE_PATH.read_bytes()

    def _make_event(self, body: dict) -> dict:
        return {
            "httpMethod": "POST",
            "body": json.dumps(body),
            "isBase64Encoded": False,
        }

    def _invoke(self, body: dict):
        """Invoke lambda_handler with fully mocked boto3."""
        mock_boto3 = MagicMock()

        table_mock = MagicMock()
        table_mock.get_item.side_effect = _dynamo_get_item
        mock_boto3.resource.return_value.Table.return_value = table_mock

        s3_mock = MagicMock()
        s3_mock.get_object.return_value = {"Body": MagicMock(
            read=lambda: self.template_bytes
        )}
        mock_boto3.client.return_value = s3_mock

        mock_botocore = MagicMock()
        mock_botocore.exceptions.ClientError = type("ClientError", (Exception,), {})

        with patch.dict(sys.modules, {
            "boto3": mock_boto3,
            "botocore": mock_botocore,
            "botocore.exceptions": mock_botocore.exceptions,
        }):
            sys.modules.pop("lambda", None)
            lm = importlib.import_module("lambda")
            return lm.lambda_handler(self._make_event(body), None)

    # ── Happy path ────────────────────────────────────────────────────────────

    def test_returns_200(self):
        res = self._invoke({"report_id": "test-001", **COVER_FIELDS})
        self.assertEqual(res["statusCode"], 200)

    def test_response_is_base64_docx(self):
        res = self._invoke({"report_id": "test-001", **COVER_FIELDS})
        self.assertTrue(res["isBase64Encoded"])
        raw = base64.b64decode(res["body"])
        # Must be a valid zip (docx is a zip)
        self.assertTrue(zipfile.is_zipfile(BytesIO(raw)), "Response is not a valid zip/docx")

    def test_cover_fields_replaced(self):
        res = self._invoke({"report_id": "test-001", **COVER_FIELDS})
        raw = base64.b64decode(res["body"])
        with zipfile.ZipFile(BytesIO(raw)) as z:
            doc_xml = z.read("word/document.xml").decode("utf-8", errors="replace")

        # Title placeholder replaced
        self.assertIn("2601-0042", doc_xml)
        self.assertIn("Auditoría de Controles TI", doc_xml)
        self.assertNotIn("26XX-XXXX", doc_xml)
        self.assertNotIn("TÍTULO DE LA AUDITORÍA", doc_xml)

        # UAI replaced
        self.assertIn("Tecnología", doc_xml)
        self.assertNotIn(">UAI de XXX<", doc_xml)

        # Recipients replaced
        self.assertIn("Juan García López", doc_xml)

    def test_headers_contain_date_and_title(self):
        res = self._invoke({"report_id": "test-001", **COVER_FIELDS})
        raw = base64.b64decode(res["body"])
        with zipfile.ZipFile(BytesIO(raw)) as z:
            h1 = z.read("word/header1.xml").decode("utf-8", errors="replace")
            h2 = z.read("word/header2.xml").decode("utf-8", errors="replace")

        # Date in header1
        self.assertIn("09/04/2026", h1)
        self.assertNotIn(">xx<", h1)

        # Title + UAI in header2
        self.assertIn("Tecnología", h2)
        self.assertIn("Auditoría de Controles TI", h2)

    def test_content_sections_present(self):
        res = self._invoke({"report_id": "test-001", **COVER_FIELDS})
        raw = base64.b64decode(res["body"])
        with zipfile.ZipFile(BytesIO(raw)) as z:
            doc_xml = z.read("word/document.xml").decode("utf-8", errors="replace")

        self.assertIn("controles de acceso lógico", doc_xml)
        self.assertIn("Gestión de accesos privilegiados", doc_xml)
        self.assertIn("madurez del control interno", doc_xml)
        self.assertIn("Plan de accesos", doc_xml)

    def test_writes_output_file(self):
        """Also writes sample_output.docx for visual inspection."""
        res = self._invoke({"report_id": "test-001", **COVER_FIELDS})
        raw = base64.b64decode(res["body"])
        OUTPUT_PATH.write_bytes(raw)
        print(f"\n✓ Output written → {OUTPUT_PATH}  ({len(raw):,} bytes)")

    # ── Validation errors ─────────────────────────────────────────────────────

    def test_missing_report_id_returns_400(self):
        res = self._invoke({**COVER_FIELDS})
        self.assertEqual(res["statusCode"], 400)
        self.assertIn("report_id", json.loads(res["body"])["error"])

    def test_missing_audit_code_returns_400(self):
        body = {"report_id": "test-001", **COVER_FIELDS}
        del body["audit_code"]
        res = self._invoke(body)
        self.assertEqual(res["statusCode"], 400)

    def test_options_returns_204(self):
        mock_boto3 = MagicMock()
        mock_botocore = MagicMock()
        mock_botocore.exceptions.ClientError = type("ClientError", (Exception,), {})
        with patch.dict(sys.modules, {
            "boto3": mock_boto3,
            "botocore": mock_botocore,
            "botocore.exceptions": mock_botocore.exceptions,
        }):
            sys.modules.pop("lambda", None)
            lm = importlib.import_module("lambda")
            res = lm.lambda_handler({"httpMethod": "OPTIONS"}, None)
        self.assertEqual(res["statusCode"], 204)


if __name__ == "__main__":
    unittest.main(verbosity=2)
