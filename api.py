import time
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


def _build_session() -> requests.Session:
    session = requests.Session()
    retry = Retry(
        total=3,
        backoff_factor=0.5,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "POST"],
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session


session = _build_session()


def _parse_json(response: requests.Response, context: str) -> dict:
    """Parse JSON safely, raising a descriptive error on failure."""
    if not response.text:
        raise ValueError(
            f"[{context}] HTTP {response.status_code} — resposta vazia (sem body)"
        )
    try:
        return response.json()
    except requests.exceptions.JSONDecodeError:
        preview = response.text[:300].replace("\n", " ")
        raise ValueError(
            f"[{context}] HTTP {response.status_code} — body não é JSON: {preview!r}"
        )


def get_product(ean: str, token: str) -> str | None:
    url = "https://gw.consultatributaria.com.br/product/api/product/GetProducts/1/18"

    headers = {
        "Authorization": token,
        "Content-Type": "application/json",
        "Origin": "https://portal.consultatributaria.com.br",
        "Referer": "https://portal.consultatributaria.com.br/",
        "User-Agent": "Mozilla/5.0",
    }

    payload = {
        "nameSearch": None,
        "ncmSearch": None,
        "eanSearch": ean,
        "userUf": "PB",
        "subsegmentoId": None,
        "subsegmentoUf": None,
        "ncmList": None,
        "cestList": None,
        "options": "1",
    }

    for attempt in range(4):
        response = session.post(url, json=payload, headers=headers, timeout=80)
        response.raise_for_status()

        if response.status_code == 204:
            # API throttling silencioso — aguarda e tenta novamente
            if attempt < 3:
                time.sleep(2 ** attempt)  # 1s, 2s, 4s
                continue
            return None  # esgotadas as tentativas

        data = _parse_json(response, f"get_product EAN={ean}")
        products = data.get("products", {}).get("data", [])
        return products[0]["id"] if products else None

    return None


_HEADERS_INFO = {
    "accept": "application/json, text/plain, */*",
    "accept-language": "pt-BR,pt;q=0.5",
    "origin": "https://portal.consultatributaria.com.br",
    "referer": "https://portal.consultatributaria.com.br/",
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-site",
    "user-agent": "Mozilla/5.0",
}


def _get_info_url(product_id: str) -> str:
    return f"https://api-k8s.consultatributaria.com.br/api_2/product/getInfoPortal/{product_id}"


def get_product_info(product_id: str, token: str) -> tuple[str, str, str]:
    headers = {**_HEADERS_INFO, "authorization": token}
    r = session.get(_get_info_url(product_id), headers=headers, timeout=80)
    r.raise_for_status()

    data = _parse_json(r, f"get_product_info id={product_id}")
    product = data.get("product", {})
    _IMG_BASE = "https://fgf-revisao-fiscal.s3.amazonaws.com/img/produtcs"

    nome = product.get("name") or ""
    ficha = data.get("descriptionNCM") or ""
    _img_path = product.get("productImage") or ""
    imagem = f"{_IMG_BASE}{_img_path}" if _img_path else ""

    return nome, ficha, imagem


def inspect_ean(ean: str, token: str) -> dict:
    """Retorna o JSON bruto completo das duas chamadas para um EAN. Uso: debug."""
    result: dict = {}

    product_id = get_product(ean, token)
    if not product_id:
        result["get_product"] = "204 — produto não encontrado ou throttling"
        return result

    result["product_id"] = product_id

    headers = {**_HEADERS_INFO, "authorization": token}
    r = session.get(_get_info_url(product_id), headers=headers, timeout=80)
    r.raise_for_status()
    result["get_product_info"] = r.json() if r.text else None

    return result
