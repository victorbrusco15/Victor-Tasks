"""
Pipeline de processamento de leads Porsche para upload no SAFE App.

Fluxo por campanha:
  input/PlanilhaB.xlsx
    → dedup interno (mesmo arquivo)
    → dedup histórico (historico/historico_leads.xlsx)
    → split Meta (utm_source=facebook) / Landing (demais)
    → output/Contatos_Meta_TIMESTAMP.xlsx
    → output/Contatos_Landing_TIMESTAMP.xlsx
    → input movido para processed/PlanilhaB_TIMESTAMP.xlsx
    → histórico atualizado
"""
import argparse
import logging
import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

import pandas as pd

BASE_PATH = Path(r"C:\Users\BRUSCOVI\OneDrive - Dr. Ing. h.c. F. Porsche AG\Documents\Python\Campaigns")
HISTORICO_DIR = "historico"
LOGS_DIR = "logs"
HISTORICO_FILE_NAME = "historico_leads.xlsx"
INPUT_FILE_NAME = "PlanilhaB.xlsx"
SOURCE_SHEET = "Leads"
META_UTM_SOURCE = "facebook"
INTERNAL_COLUMNS = ["Email_key", "CPF_key", "unique_key", "utm_source_key"]

SAFE_COLUMNS = [
    "Dealer (Must follow SAFE's App Dealer Name) *",
    "Vehicle Variant (Must follow SAFE's APP Variant Name)",
    "First Name",
    "Last Name",
    "CPF Number (Brazil)",
    "Email *",
    "Mobile No. *",
    "Compromisso de Proteção de Dados Porsche *",
    "Política de Privacidade & Consentimento (Phone) *",
    "Política de Privacidade & Consentimento (Email) *",
    "Política de Privacidade & Consentimento (Mail) *",
]

SOURCE_REQUIRED_COLUMNS = [
    "Unidade",
    "Modelo",
    "Nome",
    "CPF",
    "E-mail",
    "Telefone",
    "utm_source",
]

DEALER_MAP = {
    "Rio de Janeiro (RJ)": "Porsche Center Rio de Janeiro",
    "São Paulo - Vila Madalena (SP)": "Porsche Center São Paulo Oeste",
    "São Paulo - Itaim (SP)": "Porsche Store Itaim",
    "Barueri - Alphaville (SP)": "Porsche Center São Paulo Oeste",
    "Belo Horizonte (MG)": "Porsche Center Belo Horizonte",
    "Blumenau (SC)": "Porsche Center Blumenau",
    "Brasília (DF)": "Porsche Center Brasília",
    "Campinas (SP)": "Porsche Center Campinas",
    "Campo Grande (MS)": "Porsche Center Salvador",
    "Curitiba (PR)": "Porsche Center Curitiba",
    "Florianópolis (SC)": "Porsche Center Florianópolis",
    "Fortaleza (CE)": "Porsche Center Fortaleza",
    "Goiânia (GO)": "Porsche Center Goiânia",
    "Maringá (PR)": "Porsche Center Maringá",
    "Porto Alegre (RS)": "Porsche Center Porto Alegre",
    "Recife (PE)": "Porsche Center Recife",
    "Ribeirão Preto (SP)": "Porsche Center Ribeirão Preto",
    "Salvador (BA)": "Porsche Center Salvador",
    "São Paulo - Vila Olímpia (SP)": "Porsche Center São Paulo",
}


@dataclass
class CampaignResult:
    campaign: str
    source_rows: int
    internal_duplicates: int
    historical_duplicates: int
    meta_rows: int
    landing_rows: int
    unmapped_dealers: list[str]


def only_digits(value: object) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\D", "", str(value))


def normalize_email(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().lower()


def format_phone(phone: object) -> str:
    digits = only_digits(phone)
    if not digits:
        return ""

    # Remove DDI duplicado com prefixo 00 (ex: 0055 → 55)
    if digits.startswith("00"):
        digits = digits[2:]

    if digits.startswith("55"):
        national = digits[2:]
        if len(national) in (10, 11):
            return f"+55{national}"
        # número 55+... mas tamanho inesperado; mantém como está com +
        return f"+{digits}"

    if len(digits) in (10, 11):
        return f"+55{digits}"

    return f"+{digits}"


def normalize_name(name: object) -> str:
    """Converte para Title Case: 'FELIPE VASCONCELLOS' → 'Felipe Vasconcellos'."""
    if pd.isna(name) or not str(name).strip():
        return ""
    return str(name).strip().title()


def split_name(name: object) -> tuple[str, str]:
    normalized = normalize_name(name)
    if not normalized:
        return "", ""
    parts = normalized.split(maxsplit=1)
    return parts[0], (parts[1] if len(parts) > 1 else "")


def make_unique_key(cpf: object, email: object) -> str:
    cpf_key = only_digits(cpf)
    email_key = normalize_email(email)
    if cpf_key and email_key:
        return f"cpf:{cpf_key}|email:{email_key}"
    if cpf_key:
        return f"cpf:{cpf_key}"
    if email_key:
        return f"email:{email_key}"
    return ""


def require_columns(df: pd.DataFrame, required: Iterable[str], context: str) -> None:
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"{context}: colunas obrigatórias ausentes: {', '.join(missing)}")


def setup_logging(logs_path: Path, timestamp: str) -> None:
    logs_path.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        filename=logs_path / f"log_{timestamp}.txt",
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        force=True,
    )


def load_historico(historico_file: Path) -> pd.DataFrame:
    if not historico_file.exists():
        return pd.DataFrame(columns=SAFE_COLUMNS + ["unique_key"])

    df = pd.read_excel(historico_file, dtype=str).fillna("")
    df.columns = df.columns.str.strip()

    if "CPF Number (Brazil)" not in df.columns or "Email *" not in df.columns:
        raise ValueError(
            f"Histórico inválido em {historico_file}: esperado 'CPF Number (Brazil)' e 'Email *'"
        )

    df["CPF Number (Brazil)"] = df["CPF Number (Brazil)"].apply(only_digits)
    # Normaliza e-mail antes de montar a chave (garante case-insensitive mesmo em históricos antigos)
    df["Email *"] = df["Email *"].apply(normalize_email)
    df["unique_key"] = df.apply(
        lambda row: make_unique_key(row["CPF Number (Brazil)"], row["Email *"]),
        axis=1,
    )
    df = df[df["unique_key"] != ""].copy()
    return df.drop_duplicates(subset="unique_key", keep="first")


def build_target(df_source: pd.DataFrame) -> pd.DataFrame:
    # FIX: strip colunas ANTES de require_columns para tolerar espaços extras nos cabeçalhos
    df_source = df_source.copy()
    df_source.columns = df_source.columns.str.strip()
    df_source = df_source.fillna("")

    require_columns(df_source, SOURCE_REQUIRED_COLUMNS, "Planilha de origem")

    names = df_source["Nome"].apply(split_name)
    target = pd.DataFrame(
        {
            "Dealer (Must follow SAFE's App Dealer Name) *": (
                df_source["Unidade"].astype(str).str.strip().map(DEALER_MAP).fillna("NÃO MAPEADO")
            ),
            "Vehicle Variant (Must follow SAFE's APP Variant Name)": df_source["Modelo"].astype(str).str.strip(),
            "First Name": names.apply(lambda t: t[0]),
            "Last Name": names.apply(lambda t: t[1]),
            "CPF Number (Brazil)": df_source["CPF"].apply(only_digits),
            "Email *": df_source["E-mail"].apply(normalize_email),
            "Mobile No. *": df_source["Telefone"].apply(format_phone),
            "Compromisso de Proteção de Dados Porsche *": "Yes",
            "Política de Privacidade & Consentimento (Phone) *": "Yes",
            "Política de Privacidade & Consentimento (Email) *": "Yes",
            "Política de Privacidade & Consentimento (Mail) *": "Yes",
            # utm_source_key viaja dentro do DataFrame filtrado para evitar
            # desalinhamento de índice após deduplicação
            "utm_source_key": df_source["utm_source"].astype(str).str.strip().str.lower(),
        }
    )

    target["Email_key"] = target["Email *"]          # já normalizado acima
    target["CPF_key"] = target["CPF Number (Brazil)"]  # já apenas dígitos acima
    target["unique_key"] = target.apply(
        lambda row: make_unique_key(row["CPF_key"], row["Email_key"]),
        axis=1,
    )
    return target


def clean_export_frame(df: pd.DataFrame) -> pd.DataFrame:
    return df.drop(columns=INTERNAL_COLUMNS, errors="ignore").reindex(columns=SAFE_COLUMNS)


def save_historico(historico_file: Path, df_historico: pd.DataFrame) -> None:
    clean_export_frame(df_historico).to_excel(historico_file, index=False)


def process_campaign(
    campaign: Path,
    df_historico: pd.DataFrame,
    timestamp: str,
    dry_run: bool,
) -> tuple[pd.DataFrame, CampaignResult | None]:
    input_folder = campaign / "input"
    output_folder = campaign / "output"
    processed_folder = campaign / "processed"
    planilha = input_folder / INPUT_FILE_NAME

    if not planilha.exists():
        return df_historico, None

    output_folder.mkdir(parents=True, exist_ok=True)
    processed_folder.mkdir(parents=True, exist_ok=True)

    df_source = pd.read_excel(planilha, sheet_name=SOURCE_SHEET, dtype=str)
    df_source.columns = df_source.columns.str.strip()
    df_target = build_target(df_source)

    # Remove linhas sem nenhuma chave de identificação
    invalid_keys = df_target["unique_key"] == ""
    if invalid_keys.any():
        logging.warning(
            "%s - ignorando %s linhas sem CPF e sem e-mail", campaign.name, int(invalid_keys.sum())
        )
        df_target = df_target[~invalid_keys].copy()

    # 1) Dedup interno (mesmo arquivo)
    before_internal = len(df_target)
    df_target = df_target.drop_duplicates(subset="unique_key", keep="first").copy()
    internal_duplicates = before_internal - len(df_target)

    # 2) Dedup histórico
    historico_keys = set(df_historico.get("unique_key", pd.Series(dtype=str)).dropna())
    before_history = len(df_target)
    df_new = df_target[~df_target["unique_key"].isin(historico_keys)].copy()
    historical_duplicates = before_history - len(df_new)

    # 3) Split por canal — máscara aplicada sobre df_new (índice consistente após ambas as deduplicações)
    meta_mask = df_new["utm_source_key"].eq(META_UTM_SOURCE)
    df_meta = df_new[meta_mask].copy()
    df_landing = df_new[~meta_mask].copy()

    # Dealers não mapeados no arquivo de origem (antes de qualquer filtro)
    unmapped_dealers = sorted(
        df_source.loc[
            ~df_source["Unidade"].astype(str).str.strip().isin(DEALER_MAP), "Unidade"
        ]
        .dropna()
        .astype(str)
        .str.strip()
        .loc[lambda s: s != ""]
        .unique()
        .tolist()
    )

    if not dry_run:
        output_meta = output_folder / f"Contatos_Meta_{timestamp}.xlsx"
        output_landing = output_folder / f"Contatos_Landing_{timestamp}.xlsx"
        processed_file = processed_folder / f"PlanilhaB_{timestamp}.xlsx"

        clean_export_frame(df_meta).to_excel(output_meta, index=False)
        clean_export_frame(df_landing).to_excel(output_landing, index=False)
        shutil.move(str(planilha), str(processed_file))

    if len(df_new):
        df_historico = pd.concat([df_historico, df_new], ignore_index=True)
        df_historico = df_historico[df_historico["unique_key"] != ""].copy()
        df_historico.drop_duplicates(subset="unique_key", keep="first", inplace=True)

    result = CampaignResult(
        campaign=campaign.name,
        source_rows=len(df_source),
        internal_duplicates=internal_duplicates,
        historical_duplicates=historical_duplicates,
        meta_rows=len(df_meta),
        landing_rows=len(df_landing),
        unmapped_dealers=unmapped_dealers,
    )
    logging.info(
        "%s - origem=%s dup_interno=%s dup_historico=%s Meta=%s Landing=%s dry_run=%s",
        result.campaign,
        result.source_rows,
        result.internal_duplicates,
        result.historical_duplicates,
        result.meta_rows,
        result.landing_rows,
        dry_run,
    )
    if unmapped_dealers:
        logging.warning("%s - dealers não mapeados: %s", campaign.name, "; ".join(unmapped_dealers))

    return df_historico, result


def iter_campaigns(base_path: Path) -> Iterable[Path]:
    for item in sorted(base_path.iterdir(), key=lambda p: p.name.lower()):
        if item.is_dir() and item.name not in {HISTORICO_DIR, LOGS_DIR}:
            yield item


def run_pipeline(
    base_path: Path = BASE_PATH,
    dry_run: bool = False,
    allow_empty: bool = False,
) -> list[CampaignResult]:
    timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
    historico_path = base_path / HISTORICO_DIR
    logs_path = base_path / LOGS_DIR
    historico_file = historico_path / HISTORICO_FILE_NAME

    historico_path.mkdir(parents=True, exist_ok=True)
    setup_logging(logs_path, timestamp)

    print("Iniciando pipeline..." + (" (dry-run)" if dry_run else ""))
    logging.info("PIPELINE INICIADO dry_run=%s", dry_run)

    df_historico = load_historico(historico_file)
    initial_history_rows = len(df_historico)
    results: list[CampaignResult] = []

    try:
        for campaign in iter_campaigns(base_path):
            df_historico, result = process_campaign(campaign, df_historico, timestamp, dry_run)
            if result is None:
                continue

            results.append(result)
            print(
                f"{result.campaign}: origem={result.source_rows} | "
                f"novos Meta={result.meta_rows} | novos Landing={result.landing_rows} | "
                f"dup interno={result.internal_duplicates} | "
                f"dup histórico={result.historical_duplicates}"
            )
            if result.unmapped_dealers:
                print(f"  ⚠ Dealers não mapeados: {', '.join(result.unmapped_dealers)}")

        if not results and not allow_empty:
            raise FileNotFoundError(
                f"Nenhum {INPUT_FILE_NAME} encontrado em subpastas */input de {base_path}"
            )

        if not dry_run and results:
            save_historico(historico_file, df_historico)

        final_history_rows = len(df_historico)
        print(
            f"PIPELINE FINALIZADO. Campanhas={len(results)} | "
            f"histórico: {initial_history_rows} → {final_history_rows}"
        )
        logging.info(
            "PIPELINE FINALIZADO campanhas=%s historico=%s->%s",
            len(results),
            initial_history_rows,
            final_history_rows,
        )
    finally:
        logging.shutdown()

    return results


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Processa exports de leads Porsche para upload no SAFE App."
    )
    parser.add_argument("--base-path", type=Path, default=BASE_PATH, help="Pasta raiz Campaigns.")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Valida sem gerar arquivos, mover inputs ou atualizar histórico.",
    )
    parser.add_argument(
        "--allow-empty",
        action="store_true",
        help="Não falha quando não houver PlanilhaB.xlsx em input.",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    run_pipeline(base_path=args.base_path, dry_run=args.dry_run, allow_empty=args.allow_empty)
