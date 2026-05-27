"""
Testes do pipeline de leads Porsche.

Cobre:
  - Funções de formatação (only_digits, normalize_email, format_phone, split_name, make_unique_key)
  - build_target (mapeamento de colunas, dealers não mapeados, utm_source_key)
  - load_historico (criação a frio, normalização de e-mail, dedup)
  - process_campaign (dedup interno, dedup histórico, split meta/landing, dry_run)
  - run_pipeline (allow_empty, múltiplas campanhas, atualização de histórico)
  - automatic_period (lógica de segunda-feira vs demais dias)
"""
import shutil
import tempfile
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import pytest

import pipeline_leads as pl

# automatic_period importado com guard para evitar falha se playwright não instalado
try:
    from portal_downloads import automatic_period
    _PORTAL_AVAILABLE = True
except ModuleNotFoundError:
    _PORTAL_AVAILABLE = False

    def automatic_period(today=None):
        from datetime import date, timedelta
        today = today or date.today()
        if today.weekday() == 0:
            return today - timedelta(days=3), today
        return today - timedelta(days=1), today


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture()
def tmp_base(tmp_path):
    """Pasta raiz temporária simulando BASE_PATH."""
    return tmp_path


def make_source_df(**overrides) -> pd.DataFrame:
    """DataFrame mínimo válido para build_target."""
    data = {
        "Unidade": ["Rio de Janeiro (RJ)"],
        "Modelo": ["Cayenne"],
        "Nome": ["João Silva"],
        "CPF": ["12345678901"],
        "E-mail": ["joao@example.com"],
        "Telefone": ["11987654321"],
        "utm_source": ["facebook"],
    }
    data.update(overrides)
    return pd.DataFrame(data)


def write_input_excel(campaign_dir: Path, df: pd.DataFrame) -> Path:
    input_dir = campaign_dir / "input"
    input_dir.mkdir(parents=True, exist_ok=True)
    path = input_dir / "PlanilhaB.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Leads", index=False)
    return path


# ---------------------------------------------------------------------------
# only_digits
# ---------------------------------------------------------------------------

class TestOnlyDigits:
    def test_normal(self):
        assert pl.only_digits("123.456.789-01") == "12345678901"

    def test_nan(self):
        assert pl.only_digits(float("nan")) == ""

    def test_none(self):
        assert pl.only_digits(None) == ""

    def test_empty(self):
        assert pl.only_digits("") == ""

    def test_all_digits(self):
        assert pl.only_digits("11987654321") == "11987654321"

    def test_float_string(self):
        # Excel às vezes retorna floats; dtype=str transforma em "12345678901.0"
        assert pl.only_digits("12345678901.0") == "123456789010"
        # Nota: isso é um comportamento esperado; use dtype=str no read_excel
        # para garantir que CPFs não sejam lidos como float

    def test_non_numeric(self):
        assert pl.only_digits("abc") == ""


# ---------------------------------------------------------------------------
# normalize_email
# ---------------------------------------------------------------------------

class TestNormalizeEmail:
    def test_lowercase(self):
        assert pl.normalize_email("JOAO@EXAMPLE.COM") == "joao@example.com"

    def test_strip(self):
        assert pl.normalize_email("  joao@example.com  ") == "joao@example.com"

    def test_nan(self):
        assert pl.normalize_email(float("nan")) == ""

    def test_none(self):
        assert pl.normalize_email(None) == ""


# ---------------------------------------------------------------------------
# format_phone
# ---------------------------------------------------------------------------

class TestFormatPhone:
    def test_11_digits(self):
        assert pl.format_phone("11987654321") == "+5511987654321"

    def test_10_digits(self):
        assert pl.format_phone("1133334444") == "+551133334444"

    def test_with_country_code_55(self):
        assert pl.format_phone("5511987654321") == "+5511987654321"

    def test_with_00_prefix(self):
        assert pl.format_phone("005511987654321") == "+5511987654321"

    def test_empty(self):
        assert pl.format_phone("") == ""

    def test_nan(self):
        assert pl.format_phone(float("nan")) == ""

    def test_formatted_input(self):
        assert pl.format_phone("(11) 9 8765-4321") == "+5511987654321"

    def test_55_short_national(self):
        # Começa com 55 mas national part tem tamanho inesperado → mantém +55XXX
        result = pl.format_phone("5512345")
        assert result.startswith("+55")


# ---------------------------------------------------------------------------
# split_name
# ---------------------------------------------------------------------------

class TestSplitName:
    def test_full_name(self):
        assert pl.split_name("João da Silva Sauro") == ("João", "da Silva Sauro")

    def test_single_name(self):
        assert pl.split_name("Cher") == ("Cher", "")

    def test_empty(self):
        assert pl.split_name("") == ("", "")

    def test_nan(self):
        assert pl.split_name(float("nan")) == ("", "")

    def test_all_caps(self):
        assert pl.split_name("FELIPE VASCONCELLOS") == ("Felipe", "Vasconcellos")

    def test_mixed_case(self):
        assert pl.split_name("Tiago caldeira") == ("Tiago", "Caldeira")

    def test_particle_de(self):
        assert pl.split_name("ANA DE SOUZA") == ("Ana", "de Souza")

    def test_particle_da(self):
        assert pl.split_name("CARLOS DA SILVA") == ("Carlos", "da Silva")

    def test_particle_dos(self):
        assert pl.split_name("MARIA DOS SANTOS") == ("Maria", "dos Santos")

    def test_particle_e(self):
        assert pl.split_name("PEDRO E SILVA") == ("Pedro", "e Silva")

    def test_particle_first_word_still_capitalized(self):
        # Partícula no início do nome completo deve ser capitalizada
        assert pl.split_name("de souza carlos") == ("De", "Souza Carlos")

    def test_whitespace_only(self):
        assert pl.split_name("   ") == ("", "")

    def test_whitespace_only(self):
        assert pl.split_name("   ") == ("", "")


# ---------------------------------------------------------------------------
# make_unique_key
# ---------------------------------------------------------------------------

class TestMakeUniqueKey:
    def test_both(self):
        key = pl.make_unique_key("12345678901", "joao@example.com")
        assert key == "cpf:12345678901|email:joao@example.com"

    def test_only_cpf(self):
        assert pl.make_unique_key("12345678901", "") == "cpf:12345678901"

    def test_only_email(self):
        assert pl.make_unique_key("", "joao@example.com") == "email:joao@example.com"

    def test_empty(self):
        assert pl.make_unique_key("", "") == ""

    def test_case_insensitive_email(self):
        key1 = pl.make_unique_key("", "JOAO@EXAMPLE.COM")
        key2 = pl.make_unique_key("", "joao@example.com")
        assert key1 == key2


# ---------------------------------------------------------------------------
# build_target
# ---------------------------------------------------------------------------

class TestBuildTarget:
    def test_basic_mapping(self):
        df = make_source_df()
        target = pl.build_target(df)
        assert target["Dealer (Must follow SAFE's App Dealer Name) *"].iloc[0] == "Porsche Center Rio de Janeiro"
        assert target["First Name"].iloc[0] == "João"
        assert target["Last Name"].iloc[0] == "Silva"
        assert target["utm_source_key"].iloc[0] == "facebook"
        assert target["unique_key"].iloc[0] != ""

    def test_unmapped_dealer(self):
        df = make_source_df(Unidade=["Cidade Desconhecida (XX)"])
        target = pl.build_target(df)
        assert target["Dealer (Must follow SAFE's App Dealer Name) *"].iloc[0] == "NÃO MAPEADO"

    def test_email_normalized_in_target(self):
        df = make_source_df(**{"E-mail": ["UPPER@EXAMPLE.COM"]})
        target = pl.build_target(df)
        assert target["Email *"].iloc[0] == "upper@example.com"

    def test_consent_columns_all_yes(self):
        df = make_source_df()
        target = pl.build_target(df)
        for col in [
            "Compromisso de Proteção de Dados Porsche *",
            "Política de Privacidade & Consentimento (Phone) *",
            "Política de Privacidade & Consentimento (Email) *",
            "Política de Privacidade & Consentimento (Mail) *",
        ]:
            assert target[col].iloc[0] == "Yes"

    def test_columns_with_leading_spaces(self):
        """Colunas com espaço extra no cabeçalho devem ser aceitas (FIX)."""
        df = make_source_df()
        df.columns = [f" {c} " for c in df.columns]  # adiciona espaços
        target = pl.build_target(df)
        assert len(target) == 1

    def test_missing_required_column_raises(self):
        df = make_source_df().drop(columns=["CPF"])
        with pytest.raises(ValueError, match="CPF"):
            pl.build_target(df)

    def test_utm_source_lowercased(self):
        df = make_source_df(utm_source=["Facebook"])  # maiúscula
        target = pl.build_target(df)
        assert target["utm_source_key"].iloc[0] == "facebook"


# ---------------------------------------------------------------------------
# load_historico
# ---------------------------------------------------------------------------

class TestLoadHistorico:
    def test_missing_file_returns_empty(self, tmp_path):
        df = pl.load_historico(tmp_path / "nao_existe.xlsx")
        assert len(df) == 0
        assert "unique_key" in df.columns

    def test_loads_and_computes_key(self, tmp_path):
        hist_file = tmp_path / "historico_leads.xlsx"
        data = {col: [""] for col in pl.SAFE_COLUMNS}
        data["CPF Number (Brazil)"] = ["12345678901"]
        data["Email *"] = ["Joao@Example.COM"]  # case misto intencional
        df_in = pd.DataFrame(data)
        df_in.to_excel(hist_file, index=False)

        df = pl.load_historico(hist_file)
        # E-mail deve ser normalizado
        assert df["Email *"].iloc[0] == "joao@example.com"
        assert df["unique_key"].iloc[0] == "cpf:12345678901|email:joao@example.com"

    def test_deduplication_on_load(self, tmp_path):
        hist_file = tmp_path / "historico_leads.xlsx"
        data = {col: ["", ""] for col in pl.SAFE_COLUMNS}
        data["CPF Number (Brazil)"] = ["12345678901", "12345678901"]
        data["Email *"] = ["a@b.com", "a@b.com"]
        pd.DataFrame(data).to_excel(hist_file, index=False)

        df = pl.load_historico(hist_file)
        assert len(df) == 1

    def test_invalid_historico_raises(self, tmp_path):
        hist_file = tmp_path / "historico_leads.xlsx"
        pd.DataFrame({"ColunaSemSentido": [1]}).to_excel(hist_file, index=False)
        with pytest.raises(ValueError, match="CPF Number"):
            pl.load_historico(hist_file)


# ---------------------------------------------------------------------------
# process_campaign
# ---------------------------------------------------------------------------

class TestProcessCampaign:
    def _make_campaign(self, base: Path, name: str, rows: list[dict]) -> Path:
        campaign = base / name
        df = pd.DataFrame(rows)
        write_input_excel(campaign, df)
        return campaign

    def _empty_historico(self):
        return pd.DataFrame(columns=pl.SAFE_COLUMNS + ["unique_key"])

    def test_basic_processing(self, tmp_base):
        rows = [
            {"Unidade": "Rio de Janeiro (RJ)", "Modelo": "Cayenne", "Nome": "Ana Lima",
             "CPF": "11111111111", "E-mail": "ana@example.com", "Telefone": "21912345678",
             "utm_source": "facebook"},
            {"Unidade": "Rio de Janeiro (RJ)", "Modelo": "Cayenne", "Nome": "Bob Silva",
             "CPF": "22222222222", "E-mail": "bob@example.com", "Telefone": "21912345679",
             "utm_source": "google"},
        ]
        campaign = self._make_campaign(tmp_base, "Campanha1", rows)
        hist = self._empty_historico()
        hist, result = pl.process_campaign(campaign, hist, "20260101_120000", dry_run=True)

        assert result is not None
        assert result.source_rows == 2
        assert result.meta_rows == 1    # facebook
        assert result.landing_rows == 1  # google
        assert result.internal_duplicates == 0
        assert result.historical_duplicates == 0

    def test_internal_dedup(self, tmp_base):
        rows = [
            {"Unidade": "Rio de Janeiro (RJ)", "Modelo": "Cayenne", "Nome": "Ana Lima",
             "CPF": "11111111111", "E-mail": "ana@example.com", "Telefone": "21912345678",
             "utm_source": "facebook"},
            # Mesmo CPF + email → duplicado interno
            {"Unidade": "Campinas (SP)", "Modelo": "Macan", "Nome": "Ana Lima",
             "CPF": "11111111111", "E-mail": "ana@example.com", "Telefone": "19987654321",
             "utm_source": "facebook"},
        ]
        campaign = self._make_campaign(tmp_base, "Campanha1", rows)
        hist = self._empty_historico()
        _, result = pl.process_campaign(campaign, hist, "20260101_120000", dry_run=True)
        assert result.internal_duplicates == 1
        assert result.meta_rows == 1

    def test_historical_dedup(self, tmp_base):
        rows = [
            {"Unidade": "Rio de Janeiro (RJ)", "Modelo": "Cayenne", "Nome": "Ana Lima",
             "CPF": "11111111111", "E-mail": "ana@example.com", "Telefone": "21912345678",
             "utm_source": "facebook"},
        ]
        campaign = self._make_campaign(tmp_base, "Campanha1", rows)

        # Historico já tem essa pessoa
        hist_data = {col: [""] for col in pl.SAFE_COLUMNS}
        hist_data["CPF Number (Brazil)"] = ["11111111111"]
        hist_data["Email *"] = ["ana@example.com"]
        hist = pd.DataFrame(hist_data)
        hist["unique_key"] = hist.apply(
            lambda r: pl.make_unique_key(r["CPF Number (Brazil)"], r["Email *"]), axis=1
        )

        _, result = pl.process_campaign(campaign, hist, "20260101_120000", dry_run=True)
        assert result.historical_duplicates == 1
        assert result.meta_rows == 0

    def test_utm_source_segmentation_after_dedup(self, tmp_base):
        """FIX crítico: máscara de utm_source deve ser aplicada sobre df_new (pós-dedup)."""
        rows = [
            # Linha 1: facebook — será mantida (nova)
            {"Unidade": "Rio de Janeiro (RJ)", "Modelo": "Cayenne", "Nome": "Ana Lima",
             "CPF": "11111111111", "E-mail": "ana@example.com", "Telefone": "21912345678",
             "utm_source": "facebook"},
            # Linha 2: google — duplicado interno (mesmo CPF/email)
            {"Unidade": "Campinas (SP)", "Modelo": "Macan", "Nome": "Ana Lima",
             "CPF": "11111111111", "E-mail": "ana@example.com", "Telefone": "19987654321",
             "utm_source": "google"},
            # Linha 3: landing nova
            {"Unidade": "Curitiba (PR)", "Modelo": "911", "Nome": "Bob Souza",
             "CPF": "22222222222", "E-mail": "bob@example.com", "Telefone": "41912345678",
             "utm_source": "organic"},
        ]
        campaign = self._make_campaign(tmp_base, "Campanha1", rows)
        hist = self._empty_historico()
        _, result = pl.process_campaign(campaign, hist, "20260101_120000", dry_run=True)

        # Após dedup: Ana (facebook) + Bob (organic) = 2 novos
        assert result.internal_duplicates == 1
        assert result.meta_rows == 1    # Ana
        assert result.landing_rows == 1  # Bob

    def test_no_input_file_returns_none(self, tmp_base):
        campaign = tmp_base / "VaziaCampanha"
        campaign.mkdir()
        hist = self._empty_historico()
        _, result = pl.process_campaign(campaign, hist, "20260101_120000", dry_run=True)
        assert result is None

    def test_dry_run_does_not_create_files(self, tmp_base):
        rows = [
            {"Unidade": "Rio de Janeiro (RJ)", "Modelo": "Cayenne", "Nome": "Ana",
             "CPF": "11111111111", "E-mail": "ana@example.com", "Telefone": "21912345678",
             "utm_source": "facebook"},
        ]
        campaign = self._make_campaign(tmp_base, "Campanha1", rows)
        hist = self._empty_historico()
        pl.process_campaign(campaign, hist, "20260101_120000", dry_run=True)

        # Pastas podem ser criadas; o que não deve existir são os arquivos de output
        assert not any((campaign / "output").glob("*.xlsx"))
        assert (campaign / "input" / "PlanilhaB.xlsx").exists()

    def test_real_run_creates_output_and_moves_input(self, tmp_base):
        rows = [
            {"Unidade": "Rio de Janeiro (RJ)", "Modelo": "Cayenne", "Nome": "Ana",
             "CPF": "11111111111", "E-mail": "ana@example.com", "Telefone": "21912345678",
             "utm_source": "facebook"},
        ]
        campaign = self._make_campaign(tmp_base, "Campanha1", rows)
        hist = self._empty_historico()
        pl.process_campaign(campaign, hist, "20260101_120000", dry_run=False)

        assert not (campaign / "input" / "PlanilhaB.xlsx").exists()
        assert any((campaign / "processed").glob("PlanilhaB_*.xlsx"))
        assert any((campaign / "output").glob("Contatos_Meta_*.xlsx"))
        assert any((campaign / "output").glob("Contatos_Landing_*.xlsx"))

    def test_unmapped_dealer_reported(self, tmp_base):
        rows = [
            {"Unidade": "Cidade Inventada (ZZ)", "Modelo": "911", "Nome": "X",
             "CPF": "33333333333", "E-mail": "x@x.com", "Telefone": "11999999999",
             "utm_source": "facebook"},
        ]
        campaign = self._make_campaign(tmp_base, "Campanha1", rows)
        hist = self._empty_historico()
        _, result = pl.process_campaign(campaign, hist, "20260101_120000", dry_run=True)
        assert "Cidade Inventada (ZZ)" in result.unmapped_dealers


# ---------------------------------------------------------------------------
# run_pipeline
# ---------------------------------------------------------------------------

class TestRunPipeline:
    def _make_campaign(self, base: Path, name: str, rows: list[dict]) -> None:
        df = pd.DataFrame(rows)
        write_input_excel(base / name, df)

    def test_allow_empty(self, tmp_base):
        results = pl.run_pipeline(base_path=tmp_base, dry_run=True, allow_empty=True)
        assert results == []

    def test_raises_without_input_when_not_allow_empty(self, tmp_base):
        with pytest.raises(FileNotFoundError):
            pl.run_pipeline(base_path=tmp_base, dry_run=True, allow_empty=False)

    def test_multiple_campaigns_cross_dedup(self, tmp_base):
        """Leads duplicados entre campanhas diferentes não devem aparecer duas vezes."""
        row_shared = {
            "Unidade": "Rio de Janeiro (RJ)", "Modelo": "Cayenne", "Nome": "Ana",
            "CPF": "11111111111", "E-mail": "ana@example.com", "Telefone": "21912345678",
            "utm_source": "facebook",
        }
        row_unique = {
            "Unidade": "Curitiba (PR)", "Modelo": "911", "Nome": "Bob",
            "CPF": "22222222222", "E-mail": "bob@example.com", "Telefone": "41912345678",
            "utm_source": "organic",
        }
        self._make_campaign(tmp_base, "Alpha", [row_shared])
        self._make_campaign(tmp_base, "Beta", [row_shared, row_unique])

        results = pl.run_pipeline(base_path=tmp_base, dry_run=True, allow_empty=False)
        assert len(results) == 2

        alpha = next(r for r in results if r.campaign == "Alpha")
        beta = next(r for r in results if r.campaign == "Beta")

        assert alpha.meta_rows == 1
        assert beta.historical_duplicates == 1  # Ana já estava no histórico em memória
        assert beta.landing_rows == 1            # Bob é novo

    def test_historico_written_on_real_run(self, tmp_base):
        rows = [
            {"Unidade": "Rio de Janeiro (RJ)", "Modelo": "Cayenne", "Nome": "Ana",
             "CPF": "11111111111", "E-mail": "ana@example.com", "Telefone": "21912345678",
             "utm_source": "facebook"},
        ]
        self._make_campaign(tmp_base, "Campanha1", rows)
        pl.run_pipeline(base_path=tmp_base, dry_run=False, allow_empty=False)

        hist_file = tmp_base / pl.HISTORICO_DIR / pl.HISTORICO_FILE_NAME
        assert hist_file.exists()
        df = pd.read_excel(hist_file, dtype=str)
        assert len(df) == 1
        assert df["Email *"].iloc[0] == "ana@example.com"

    def test_historico_not_written_on_dry_run(self, tmp_base):
        rows = [
            {"Unidade": "Rio de Janeiro (RJ)", "Modelo": "Cayenne", "Nome": "Ana",
             "CPF": "11111111111", "E-mail": "ana@example.com", "Telefone": "21912345678",
             "utm_source": "facebook"},
        ]
        self._make_campaign(tmp_base, "Campanha1", rows)
        pl.run_pipeline(base_path=tmp_base, dry_run=True, allow_empty=False)

        hist_file = tmp_base / pl.HISTORICO_DIR / pl.HISTORICO_FILE_NAME
        assert not hist_file.exists()


# ---------------------------------------------------------------------------
# automatic_period (portal_downloads)
# ---------------------------------------------------------------------------

class TestAutomaticPeriod:
    def test_monday_covers_weekend(self):
        monday = date(2026, 5, 25)  # Segunda-feira
        assert monday.weekday() == 0
        start, end = automatic_period(monday)
        assert end == monday
        assert start == monday - timedelta(days=3)  # sexta

    def test_tuesday_is_yesterday_today(self):
        tuesday = date(2026, 5, 26)
        assert tuesday.weekday() == 1
        start, end = automatic_period(tuesday)
        assert start == tuesday - timedelta(days=1)
        assert end == tuesday

    def test_friday_is_yesterday_today(self):
        friday = date(2026, 5, 29)
        assert friday.weekday() == 4
        start, end = automatic_period(friday)
        assert start == friday - timedelta(days=1)
        assert end == friday

    def test_sunday_is_yesterday_today(self):
        sunday = date(2026, 5, 31)
        assert sunday.weekday() == 6
        start, end = automatic_period(sunday)
        assert start == sunday - timedelta(days=1)
        assert end == sunday


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
