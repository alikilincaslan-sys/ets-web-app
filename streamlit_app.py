import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime

from io import BytesIO

from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList

from ets_model import ets_hesapla
from data_cleaning import clean_ets_input, filter_intensity_outliers_by_fuel


# -------------------------
# Default values (V001 Stable)
# -------------------------
DEFAULTS = {
    "price_range": (5, 20),
    "agk": 1.00,
    "benchmark_top_pct": 100,
    "price_method": "Market Clearing",
    "slope_bid": 150,
    "slope_ask": 150,
    "spread": 1.0,
    "do_clean": False,
    "lower_pct": 1.0,
    "upper_pct": 2.0,
}

st.set_page_config(page_title="ETS Geliştirme Modülü V001", layout="wide")

st.title("ETS Geliştirme Modülü V001")

# -------------------------
# Model açıklaması (tek blok - düzeltilmiş)
# -------------------------
with st.expander("📌 Model Açıklaması / Sliderlar neyi değiştiriyor?", expanded=True):
    st.markdown(
        """
### ETS Geliştirme Modülü – Model Açıklaması

Bu arayüz, elektrik üretim sektörüne yönelik **tesis bazlı** ve **piyasa tutarlı** bir **ETS (Emisyon Ticaret Sistemi)** simülasyonu oluşturur.

**Veri girişi**
- Excel’deki **tüm sekmeleri** okur ve birleştirir (**FuelType = sekme adı**).
- Beklenen kolonlar: `Plant`, `Generation_MWh`, `Emissions_tCO2`

**Benchmark (yakıt bazlı)**
- Yakıt türü içinde üretim ağırlıklı benchmark hesaplanır.
- **Benchmark Top %**: Yakıt içindeki “en düşük intensity” dilimini seçer:
  - 100% = tüm tesisler (varsayılan)
  - 10% / 20% = en temiz dilim (daha sıkı benchmark)

**AGK (Adil Geçiş Katsayısı)**
- Tahsis yoğunluğu formülü:
  - **Tᵢ = Iᵢ + AGK × (B_fuel − Iᵢ)**
- AGK=1 → Benchmark’a tam yaklaşır (varsayılan)
- AGK=0 → Tesis kendi yoğunluğunda kalır

**Karbon fiyatı (tek piyasa)**
- Tüm tesisler tek piyasada birleşir ve **tek karbon fiyatı** oluşur.
- **Price Method**
  - Market Clearing: arz-talep kesişimi
  - ACC: alıcıların p_bid değerlerinin (net yükümlülükle ağırlıklı) ortalaması
- **Carbon Price Range (min–max)**: fiyat bu aralıkta kalır.

**Market Calibration**
- β_bid: alıcıların fiyat hassasiyeti
- β_ask: satıcıların fiyat hassasiyeti
- Spread: BID/ASK ayrışması için ek fark

**Veri Temizleme (opsiyonel)**
- Cleaning OFF ise sadece temel temizlik yapılır.
- Cleaning ON ise intensity outlier’lar benchmark bandına göre filtrelenir:
  - lo = B × (1 − L)
  - hi = B × (1 + U)

**Çıktılar**
- Sonuç tabloları + Excel rapor (çok sayfalı) + grafikler (Supply–Demand ve Top-20 cashflow)
"""
    )

# -------------------------
# Sidebar: Reset
# -------------------------
st.sidebar.header("Model Parameters")

if st.sidebar.button("🔄 Reset to Default"):
    st.session_state["price_range"] = DEFAULTS["price_range"]
    st.session_state["agk"] = DEFAULTS["agk"]
    st.session_state["benchmark_top_pct"] = DEFAULTS["benchmark_top_pct"]
    st.session_state["price_method"] = DEFAULTS["price_method"]
    st.session_state["slope_bid"] = DEFAULTS["slope_bid"]
    st.session_state["slope_ask"] = DEFAULTS["slope_ask"]
    st.session_state["spread"] = DEFAULTS["spread"]
    st.session_state["do_clean"] = DEFAULTS["do_clean"]
    st.session_state["lower_pct"] = DEFAULTS["lower_pct"]
    st.session_state["upper_pct"] = DEFAULTS["upper_pct"]
    st.rerun()

# -------------------------
# Sidebar: sliders (session_state bağlı)
# -------------------------
price_min, price_max = st.sidebar.slider(
    "Carbon Price Range (€/tCO₂)",
    min_value=0,
    max_value=200,
    value=st.session_state.get("price_range", DEFAULTS["price_range"]),
    step=1,
    key="price_range",
    help="Clearing price bu aralık içinde bulunur.",
)
st.sidebar.caption("Default: (5, 20)")

agk = st.sidebar.slider(
    "Just Transition Coefficient (AGK)",
    min_value=0.0,
    max_value=1.0,
    value=float(st.session_state.get("agk", DEFAULTS["agk"])),
    step=0.05,
    key="agk",
    help="AGK=1→Benchmark, AGK=0→Tesis yoğunluğu. Tᵢ = Iᵢ + AGK×(B − Iᵢ)",
)
st.sidebar.caption("Default: AGK = 1.00")

st.sidebar.subheader("Benchmark Settings")
benchmark_top_pct = st.sidebar.select_slider(
    "Benchmark = Best plants (by intensity) %",
    options=[10, 20, 30, 40, 50, 60, 70, 80, 90, 100],
    value=int(st.session_state.get("benchmark_top_pct", DEFAULTS["benchmark_top_pct"])),
    key="benchmark_top_pct",
    help="Yakıt bazında benchmark, intensity düşük olan en iyi dilimden hesaplanır. 100=tüm tesisler.",
)
st.sidebar.caption("Default: 100")

st.sidebar.subheader("Carbon Price Method")
_methods = ["Market Clearing", "Average Compliance Cost"]
_default_method = st.session_state.get("price_method", DEFAULTS["price_method"])
if _default_method not in _methods:
    _default_method = "Market Clearing"

price_method = st.sidebar.selectbox(
    "Price calculation method",
    options=_methods,
    index=_methods.index(_default_method),
    key="price_method",
    help="Market Clearing: arz-talep kesişimi. ACC: alıcıların p_bid (net_ets ile ağırlıklı) ortalaması.",
)
st.sidebar.caption("Default: Market Clearing")

st.sidebar.subheader("Market Calibration")

slope_bid = st.sidebar.slider(
    "Bid Slope (β_bid)",
    min_value=10,
    max_value=500,
    value=int(st.session_state.get("slope_bid", DEFAULTS["slope_bid"])),
    step=10,
    key="slope_bid",
    help="Alıcıların (kirli tesis) ödeme isteği hassasiyeti.",
)
st.sidebar.caption("Default: 150")

slope_ask = st.sidebar.slider(
    "Ask Slope (β_ask)",
    min_value=10,
    max_value=500,
    value=int(st.session_state.get("slope_ask", DEFAULTS["slope_ask"])),
    step=10,
    key="slope_ask",
    help="Satıcıların (temiz tesis) satış isteği hassasiyeti.",
)
st.sidebar.caption("Default: 150")

spread = st.sidebar.slider(
    "Bid/Ask Spread (€/tCO₂)",
    min_value=0.0,
    max_value=10.0,
    value=float(st.session_state.get("spread", DEFAULTS["spread"])),
    step=0.5,
    key="spread",
    help="Spread eklemek bid/ask aynı görünmesini azaltır.",
)
st.sidebar.caption("Default: 1.0")

# FX rate for TL conversion (used in briefing note)
fx_rate = st.sidebar.number_input(
    "FX Rate (TL/€)",
    min_value=0.0,
    value=float(st.session_state.get("fx_rate", 35.0)),
    step=0.5,
    key="fx_rate",
    help="Bilgi notunda €/MWh değerlerini TL/MWh'ye çevirmek için kullanılır.",
)

st.sidebar.divider()
st.sidebar.caption("Excel'de beklenen kolonlar: Plant, Generation_MWh, Emissions_tCO2")
st.sidebar.caption("Sekme adı FuelType olarak alınır.")

# -------------------------
# Data Cleaning Controls
# -------------------------
st.sidebar.subheader("Data Cleaning")

do_clean = st.sidebar.toggle(
    "Apply cleaning rules?",
    value=bool(st.session_state.get("do_clean", DEFAULTS["do_clean"])),
    key="do_clean",
    help="Kapalıysa (Hayır), outlier filtresi uygulanmaz.",
)
st.sidebar.caption("Default: OFF")

lower_pct = st.sidebar.slider(
    "Lower bound vs Benchmark (L)",
    min_value=0.0,
    max_value=1.0,
    value=float(st.session_state.get("lower_pct", DEFAULTS["lower_pct"])),
    step=0.05,
    key="lower_pct",
    help="lo = B*(1-L). L=1.0 => lo=0. L=0.5 => lo=0.5B.",
)
st.sidebar.caption("Default: 1.0")

upper_pct = st.sidebar.slider(
    "Upper bound vs Benchmark (U)",
    min_value=0.0,
    max_value=2.0,
    value=float(st.session_state.get("upper_pct", DEFAULTS["upper_pct"])),
    step=0.05,
    key="upper_pct",
    help="hi = B*(1+U). U=1.0 => hi=2B. U=2.0 => hi=3B.",
)
st.sidebar.caption("Default: 2.0")

# -------------------------
# Excel upload
# -------------------------
uploaded = st.file_uploader("Excel veri dosyasını yükleyin (.xlsx)", type=["xlsx"])


def read_all_sheets(file) -> pd.DataFrame:
    xls = pd.ExcelFile(file)
    frames = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        df["FuelType"] = sheet
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


def build_market_curve(sonuc_df: pd.DataFrame, price_min: int, price_max: int, step: int = 1) -> pd.DataFrame:
    prices = np.arange(price_min, price_max + step, step)

    buyers = sonuc_df[sonuc_df["net_ets"] > 0][["net_ets", "p_bid"]].copy()
    sellers = sonuc_df[sonuc_df["net_ets"] < 0][["net_ets", "p_ask"]].copy()

    rows = []
    for p in prices:
        if not buyers.empty:
            q0 = buyers["net_ets"].to_numpy()
            p_bid_arr = buyers["p_bid"].to_numpy()
            denom = np.maximum(p_bid_arr - price_min, 1e-6)
            frac = 1.0 - (p - price_min) / denom
            demand = float(np.sum(q0 * np.clip(frac, 0.0, 1.0)))
        else:
            demand = 0.0

        if not sellers.empty:
            q0 = (-sellers["net_ets"]).to_numpy()
            p_ask_arr = sellers["p_ask"].to_numpy()
            denom = np.maximum(price_max - p_ask_arr, 1e-6)
            frac = (p - p_ask_arr) / denom
            supply = float(np.sum(q0 * np.clip(frac, 0.0, 1.0)))
        else:
            supply = 0.0

        rows.append({"Price": float(p), "Total_Demand": demand, "Total_Supply": supply})

    return pd.DataFrame(rows)



# -------------------------
# Briefing Note (Word) Helpers
# -------------------------
def _safe_float(x, default=0.0):
    try:
        if x is None or (isinstance(x, float) and np.isnan(x)):
            return default
        return float(x)
    except Exception:
        return default


def build_tl_mwh_chart_png(sonuc_df: pd.DataFrame, fx_rate: float) -> BytesIO:
    """Create a simple bar chart (horizontal) for Net ETS impact (TL/MWh) across all plants."""
    dfc = sonuc_df.copy()
    # Net cashflow per MWh is the most consistent 'impact' indicator (can be negative for net revenue).
    if "ets_net_cashflow_€/MWh" not in dfc.columns:
        # fallback: try to derive from total and generation
        if "ets_net_cashflow_€" in dfc.columns and "Generation_MWh" in dfc.columns:
            dfc["ets_net_cashflow_€/MWh"] = dfc["ets_net_cashflow_€"] / dfc["Generation_MWh"].replace(0, np.nan)
        else:
            dfc["ets_net_cashflow_€/MWh"] = np.nan

    dfc["ets_net_cashflow_TL/MWh"] = dfc["ets_net_cashflow_€/MWh"] * float(fx_rate)

    # Sort low -> high
    dfc = dfc.sort_values("ets_net_cashflow_TL/MWh", ascending=True).reset_index(drop=True)

    # Keep labels readable: if too many plants, still plot but increase height
    n = len(dfc)
    fig_h = max(6.0, min(40.0, 0.25 * n))
    fig, ax = plt.subplots(figsize=(11, fig_h))
    ax.barh(dfc["Plant"], dfc["ets_net_cashflow_TL/MWh"])
    ax.set_xlabel("Net ETS Etkisi (TL/MWh)")
    ax.set_ylabel("Santral")
    ax.set_title("Santral Bazlı Net ETS Etkisi (TL/MWh)\n(Düşükten yükseğe sıralı)")
    ax.grid(True, axis="x", alpha=0.3)

    buf = BytesIO()
    fig.tight_layout()
    fig.savefig(buf, format="png", dpi=200)
    plt.close(fig)
    buf.seek(0)
    return buf


def generate_briefing_note_docx(
    sonuc_df: pd.DataFrame,
    benchmark_map: dict,
    clearing_price: float,
    price_method: str,
    price_min: float,
    price_max: float,
    agk: float,
    benchmark_top_pct: int,
    slope_bid: float,
    slope_ask: float,
    spread: float,
    do_clean: bool,
    lower_pct: float,
    upper_pct: float,
    df_all_raw: pd.DataFrame,
    df_all_used: pd.DataFrame,
    removed_df: pd.DataFrame,
    fx_rate: float,
    soma_example: dict | None = None,
) -> BytesIO:
    """Build a client-ready briefing note in Turkish as DOCX. Values are inserted from model outputs."""
    total_gen_mwh = _safe_float(df_all_used.get("Generation_MWh", pd.Series(dtype=float)).sum(), 0.0)
    total_emis_t = _safe_float(df_all_used.get("Emissions_tCO2", pd.Series(dtype=float)).sum(), 0.0)

    # Benchmarks table
    bench_rows = sorted([(k, _safe_float(v, np.nan)) for k, v in benchmark_map.items()], key=lambda x: str(x[0]))

    # Key metrics
    total_cost_eur = _safe_float(sonuc_df.get("ets_cost_total_€", pd.Series(dtype=float)).sum(), 0.0)
    total_rev_eur = _safe_float(sonuc_df.get("ets_revenue_total_€", pd.Series(dtype=float)).sum(), 0.0)
    net_cf_eur = _safe_float(sonuc_df.get("ets_net_cashflow_€", pd.Series(dtype=float)).sum(), 0.0)

    doc = Document()

    # Title
    title = doc.add_paragraph("Elektrik Üretim Sektörü için Emisyon Ticaret Sistemi (ETS)")
    title.runs[0].bold = True
    title.runs[0].font.size = Pt(16)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = doc.add_paragraph("Benchmark ve Karbon Fiyatı Hesaplama Modülü – Bilgi Notu")
    subtitle.runs[0].italic = True
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(f"Tarih: {datetime.now().strftime('%d.%m.%Y')}").alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("")

    # 1. Scope
    doc.add_heading("1. Çalışmanın Kapsamı ve Amacı", level=2)
    doc.add_paragraph(
        "Bu çalışma, 2024 yılına ait gerçekleşmiş elektrik üretimi ve emisyon verileri esas alınarak, "
        "2026–2027 döneminde uygulanması öngörülen Emisyon Ticaret Sistemi (ETS) kapsamında elektrik üretim santrallerinin "
        "karşılaşabileceği karbon maliyetlerinin analiz edilmesi amacıyla geliştirilmiştir."
    )
    doc.add_paragraph(
        "Çalışmanın temel amacı, geçmiş yıl verilerini referans alarak orta vadeli ETS uygulama dönemine yönelik karbon fiyatı "
        "ve maliyet etkilerini öngören; adil, piyasa temelli ve uygulanabilir bir analiz çerçevesi sunmaktır."
    )

    # 2. Coverage & benchmark
    doc.add_heading("2. ETS Kapsamı ve Benchmark Yaklaşımı", level=2)
    doc.add_paragraph(
        "Bu çalışmada benchmark hesaplamaları yakıt bazlı olarak gerçekleştirilmiştir. Elektrik üretim santralleri, kullandıkları yakıt türüne göre ayrıştırılmış "
        "ve her yakıt grubu için ayrı benchmark (referans emisyon yoğunluğu) değerleri hesaplanmıştır."
    )
    doc.add_paragraph(
        "Bununla birlikte, ETS piyasası elektrik üretim sektörü açısından bütüncül olarak ele alınmış; karbon fiyatı hesaplamasında tüm elektrik üretim santralleri tek bir piyasada değerlendirilmiştir."
    )
    doc.add_paragraph(
        "Not: Bu çalışmada, SKDM kapsamındaki diğer sanayi sektörleri, ilgili dönem için detaylı ve karşılaştırılabilir veri bulunmaması nedeniyle ETS piyasasına dahil edilmemiştir. Analiz yalnızca elektrik üretim sektörü ile sınırlandırılmıştır."
    )

    # 3. Benchmarks
    doc.add_heading("3. Benchmark Yapısı ve Yakıt Bazlı Değerler", level=2)
    p = doc.add_paragraph(
        "2024 yılı gerçekleşmiş üretim ve emisyon verilerine dayalı olarak hesaplanan yakıt bazlı benchmark emisyon yoğunlukları (tCO2/MWh) aşağıda sunulmaktadır:"
    )
    table = doc.add_table(rows=1, cols=2)
    hdr = table.rows[0].cells
    hdr[0].text = "Yakıt Türü"
    hdr[1].text = "Benchmark (tCO2/MWh)"
    for ft, b in bench_rows:
        row = table.add_row().cells
        row[0].text = str(ft)
        row[1].text = f"{b:.4f}" if not np.isnan(b) else "N/A"

    doc.add_paragraph(
        "Bu benchmark değerleri, 2026–2027 ETS uygulama döneminde tahsisat hesaplamalarında referans olarak kullanılmaktadır."
    )

    # 4. Reference year stats
    doc.add_heading("4. Referans Yıl Üretim ve Emisyon Profili (2024)", level=2)
    doc.add_paragraph(
        f"Model kapsamında değerlendirilen ETS’ye tabi elektrik üretim santralleri, 2024 yılında toplam {total_gen_mwh:,.0f} MWh elektrik üretimi gerçekleştirmiştir. "
        f"Aynı dönemde bu santrallerden kaynaklanan toplam karbondioksit (CO2) emisyonu {total_emis_t/1e6:,.2f} milyon ton olarak hesaplanmıştır."
    )

    # 5. Allocation - production weighted + AGK
    doc.add_heading("5. Tahsisat Hesaplama Yöntemi (Üretim-Ağırlıklı Benchmark)", level=2)
    doc.add_paragraph(
        "Bu çalışmada tahsis edilen emisyon miktarları, santral bazında üretim-ağırlıklı benchmark yaklaşımı kullanılarak hesaplanmıştır. "
        "Bu yaklaşımda, her bir santral için tahsis edilen emisyon miktarı, santralin elektrik üretim miktarı ile yakıt türüne özgü benchmark emisyon yoğunluğunun çarpımı yoluyla belirlenmektedir."
    )
    doc.add_paragraph("Tahsis Edilen Emisyon (tCO2) = Elektrik Üretimi (MWh) × Yakıt Bazlı Benchmark (tCO2/MWh)")

    doc.add_heading("6. AGK (α) Katsayısının Tahsisat Hesaplamalarındaki Rolü", level=2)
    doc.add_paragraph(
        "Modelde, üretim-ağırlıklı benchmark yaklaşımına ek olarak AGK (α) katsayısı uygulanmıştır. AGK katsayısı, benchmark değerlerinin geçiş dönemi boyunca kademeli ve kontrollü şekilde ayarlanmasını sağlayan bir yumuşatma parametresidir."
    )
    doc.add_paragraph(
        "AGK uygulanması durumunda tahsisat hesaplaması: Tahsis Edilen Emisyon (tCO2) = Elektrik Üretimi (MWh) × Yakıt Bazlı Benchmark (tCO2/MWh) × AGK (α)"
    )

    # 7. Net obligation
    doc.add_heading("7. Net ETS Yükümlülüğünün Hesaplanması", level=2)
    doc.add_paragraph("Net ETS Yükümlülüğü (tCO2) = Gerçekleşen Emisyon – Tahsis Edilen Emisyon")
    doc.add_paragraph(
        "Pozitif değerler, santralin ETS kapsamında piyasadan ilave emisyon izni satın alması gerektiğini; negatif değerler ise santralin emisyon fazlası bulunduğunu ve piyasaya arz sağlayabileceğini ifade etmektedir."
    )

    # 8. Carbon price method
    doc.add_heading("8. Karbon Fiyatı Hesaplama Yöntemi (2026–2027 Dönemi)", level=2)
    doc.add_paragraph(
        "Karbon fiyatı, 2026–2027 döneminde ETS’nin yürürlükte olduğu varsayımı altında, arz-talep temelli piyasa dengeleme (market clearing) yaklaşımı kullanılarak hesaplanmıştır."
    )
    doc.add_paragraph(
        f"Bu yöntem sonucunda, 2026–2027 dönemi için karbon fiyatı {clearing_price:.2f} €/tCO2 olarak hesaplanmıştır (yöntem: {price_method}; fiyat aralığı: {price_min}–{price_max} €/tCO2)."
    )

    # 9. Just transition, security of supply, AGK rationale
    doc.add_heading("9. Adil Geçiş, Arz Güvenliği ve AGK Katsayısının Önemi", level=2)
    doc.add_paragraph(
        "Model, iklim değişikliğiyle mücadele hedeflerini desteklerken, elektrik arz güvenliği, ekonomik sürdürülebilirlik ve sosyal etkiler açısından bütüncül bir yaklaşım benimsemektedir. "
        "Türkiye elektrik sistemi açısından kömür santralleri, geçiş döneminde baz yük üretimi ve sistem güvenliği bakımından hâlen önemli bir rol oynamaktadır."
    )
    doc.add_paragraph(
        "Mevcut benchmark sisteminde, AGK (Adil Geçiş Katsayısı) uygulanmadan yapılan tahsisat hesaplamaları, özellikle teknolojik olarak daha eski ve emisyon yoğunluğu yüksek kömür santrallerinin orantısız biçimde yüksek karbon maliyetleriyle karşı karşıya kalmasına yol açabilmektedir. "
        "Buna karşılık, aynı yakıtı kullanmasına rağmen daha yeni teknolojiye sahip ve görece düşük emisyon yoğunluğu bulunan santraller, benchmark sistemi içerisinde orantısız biçimde avantajlı konuma geçebilmektedir."
    )
    doc.add_paragraph(
        "Bu durum, aşırı ceza ve aşırı ödül mekanizmalarının oluşmasına neden olmakta ve daha dengeli, öngörülebilir ve nominal bir piyasa yapısını zayıflatabilmektedir. "
        "Bu çerçevede modelde kullanılan AGK (α) katsayısı (adil geçiş katsayısı), söz konusu uç etkileri yumuşatmayı; aşırı cezalandırma ve aşırı ödüllendirme davranışlarını sınırlayarak geçiş süreciyle uyumlu bir karbon piyasası oluşmasını sağlamayı amaçlamaktadır."
    )

    if soma_example:
        doc.add_paragraph(
            f"Örnek (Soma B): AGK=1.00 varsayımı altında yıllık emisyon maliyeti {soma_example.get('agk1_cost_eur', 'N/A')} €, "
            f"AGK={soma_example.get('agk_sel', agk):.2f} varsayımı altında {soma_example.get('agk_sel_cost_eur', 'N/A')} € olarak hesaplanmıştır."
        )

    # 10. €/MWh and TL/MWh
    doc.add_heading("10. Santral Bazlı Elektrik Üretimi Başına Karbon Maliyeti (€/MWh ve TL/MWh)", level=2)
    doc.add_paragraph(
        "ETS’nin elektrik üretim maliyetleri üzerindeki etkisini daha açık ve karşılaştırılabilir biçimde ortaya koymak amacıyla, santral bazında birim elektrik üretimi başına karbon maliyeti hesaplanmıştır. Bu maliyet göstergesi hem €/MWh hem de TL/MWh cinsinden sunulmaktadır."
    )
    doc.add_paragraph(
        "Karbon Maliyeti (€/MWh) = Net ETS Yükümlülüğü (tCO2) × Karbon Fiyatı (€/tCO2) ÷ Elektrik Üretimi (MWh)"
    )
    doc.add_paragraph(
        f"Karbon Maliyeti (TL/MWh) = Karbon Maliyeti (€/MWh) × Döviz Kuru (TL/€). Bu bilgi notunda kullanılan dönüşüm kuru: {float(fx_rate):.2f} TL/€."
    )

    # 11. Chart
    doc.add_heading("11. Grafiksel Gösterim: Santral Bazlı Net ETS Etkisi (TL/MWh)", level=2)
    doc.add_paragraph(
        "ETS’nin santral bazında birim elektrik üretim maliyeti üzerindeki etkisini sade ve karşılaştırılabilir biçimde göstermek amacıyla, tüm santraller için TL/MWh cinsinden net ETS etkisi bir sütun grafikte sunulmuştur. Santraller düşükten yükseğe doğru sıralanmıştır."
    )
    chart_png = build_tl_mwh_chart_png(sonuc_df, fx_rate=float(fx_rate))
    doc.add_picture(chart_png, width=Inches(6.5))
    cap = doc.add_paragraph("Şekil 1. Santral bazlı net ETS etkisi (TL/MWh) – düşükten yükseğe sıralı.")
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 12. Assumptions & sliders (concise)
    doc.add_heading("12. Varsayımlar ve Model Parametreleri", level=2)
    doc.add_paragraph(
        "Bu bölümde, modelin şeffaflığı ve senaryo karşılaştırmalarının tutarlılığı açısından temel varsayımlar ve arayüz parametreleri özetlenmektedir."
    )

    # Bullet-like paragraphs (official tone)
    items = [
        f"Referans veri yılı: 2024 (üretim ve emisyon gerçekleşmeleri). Hesaplamalar 2026–2027 ETS dönemi varsayımı altında yapılmıştır.",
        f"Karbon fiyatı yöntemi: {price_method}. Fiyat aralığı: {price_min}–{price_max} €/tCO2.",
        f"AGK (α): {agk:.2f}. Benchmark Top %: {benchmark_top_pct}. Yakıt bazlı benchmark yaklaşımı uygulanmıştır.",
        f"Piyasa kalibrasyonu: β_bid={slope_bid}, β_ask={slope_ask}, spread={spread}.",
        f"Veri temizleme: {'Açık' if do_clean else 'Kapalı'}." + (f" Outlier bandı: [{1-lower_pct:.2f}B, {1+upper_pct:.2f}B]." if do_clean else ""),
        f"Kur varsayımı (TL/€): {float(fx_rate):.2f}.",
        f"Toplam ETS maliyeti: {total_cost_eur:,.0f} €, toplam ETS geliri: {total_rev_eur:,.0f} €, net nakit akışı: {net_cf_eur:,.0f} €.",
    ]
    for it in items:
        para = doc.add_paragraph(it, style=None)
        para.paragraph_format.space_after = Pt(4)

    # Save to bytes
    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out

if uploaded is None:
    st.info("Lütfen bir Excel yükleyin.")
    st.stop()

try:
    df_all_raw = read_all_sheets(uploaded)
except Exception as e:
    st.error(f"Excel okunurken hata oluştu: {e}")
    st.stop()

st.subheader("Yüklenen veri (ham / birleştirilmiş)")
st.dataframe(df_all_raw.head(50), use_container_width=True)

# -------------------------
# Cleaning
# -------------------------
st.subheader("Veri Temizleme (opsiyonel)")

df_all = df_all_raw.copy()

try:
    df_all = clean_ets_input(df_all)
except Exception as e:
    st.error(f"Temel temizlikte hata: {e}")
    st.stop()

removed_df = pd.DataFrame()

if do_clean:
    before = len(df_all)
    try:
        df_all, removed_df = filter_intensity_outliers_by_fuel(
            df_all, lower_pct=lower_pct, upper_pct=upper_pct
        )
    except Exception as e:
        st.error(f"Outlier filtresinde hata: {e}")
        st.stop()

    after = len(df_all)
    st.info(
        f"Outlier filtresi: {before - after} satır çıkarıldı "
        f"({before:,} → {after:,}). Band: [{1-lower_pct:.2f}B, {1+upper_pct:.2f}B]"
    )
    if not removed_df.empty:
        with st.expander("Çıkarılan outlier satırlar (önizleme)"):
            st.dataframe(removed_df.head(200), use_container_width=True)
else:
    st.warning("Temizleme kapalı: (sadece temel temizlik yapıldı)")

st.subheader("Modelde kullanılacak veri (ilk 50 satır)")
st.dataframe(df_all.head(50), use_container_width=True)

# -------------------------
# Run model
# -------------------------
if st.button("Run ETS Model"):
    try:
        sonuc_df, benchmark_map, clearing_price = ets_hesapla(
            df_all,
            price_min,
            price_max,
            agk,
            slope_bid=slope_bid,
            slope_ask=slope_ask,
            spread=spread,
            benchmark_top_pct=int(benchmark_top_pct),
            price_method=price_method,
        )

        st.success(f"Carbon Price ({price_method}): {clearing_price:.2f} €/tCO₂")
        st.caption(f"Benchmark method: Best {benchmark_top_pct}% (production-share, by lowest intensity)")

        st.subheader("Benchmark (yakıt bazında)")
        bench_df = (
            pd.DataFrame([{"FuelType": k, "Benchmark_B_fuel": v} for k, v in benchmark_map.items()])
            .sort_values("FuelType")
            .reset_index(drop=True)
        )
        st.dataframe(bench_df, use_container_width=True)

        total_cost = float(sonuc_df["ets_cost_total_€"].sum())
        total_revenue = float(sonuc_df["ets_revenue_total_€"].sum())
        net_cashflow = float(sonuc_df["ets_net_cashflow_€"].sum())

        c1, c2, c3 = st.columns(3)
        c1.metric("Toplam ETS Maliyeti (€)", f"{total_cost:,.0f}")
        c2.metric("Toplam ETS Geliri (€)", f"{total_revenue:,.0f}")
        c3.metric("Net Nakit Akışı (€)", f"{net_cashflow:,.0f}")

        st.subheader("ETS Sonuçları – Alıcılar (Net ETS > 0)")
        buyers_df = sonuc_df[sonuc_df["net_ets"] > 0].copy()
        st.dataframe(
            buyers_df[
                [
                    "Plant",
                    "FuelType",
                    "net_ets",
                    "carbon_price",
                    "ets_cost_total_€",
                    "ets_cost_€/MWh",
                    "ets_net_cashflow_€",
                    "ets_net_cashflow_€/MWh",
                ]
            ],
            use_container_width=True,
        )

        st.subheader("ETS Sonuçları – Satıcılar (Net ETS < 0)")
        sellers_df = sonuc_df[sonuc_df["net_ets"] < 0].copy()
        st.dataframe(
            sellers_df[
                [
                    "Plant",
                    "FuelType",
                    "net_ets",
                    "carbon_price",
                    "ets_revenue_total_€",
                    "ets_revenue_€/MWh",
                    "ets_net_cashflow_€",
                    "ets_net_cashflow_€/MWh",
                ]
            ],
            use_container_width=True,
        )

        st.subheader("Tüm Sonuçlar (ham tablo)")
        st.dataframe(sonuc_df, use_container_width=True)

        curve_df = build_market_curve(sonuc_df, price_min, price_max, step=1)

        cashflow_top20 = (
            sonuc_df[["Plant", "FuelType", "ets_net_cashflow_€"]]
            .copy()
            .sort_values("ets_net_cashflow_€", ascending=False)
            .head(20)
        )

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            summary_df = pd.DataFrame(
                {
                    "Metric": [
                        "Carbon Price (€/tCO₂)",
                        "Price Method",
                        "Total ETS Cost (€)",
                        "Total ETS Revenue (€)",
                        "Net Cashflow (€)",
                        "Price Min",
                        "Price Max",
                        "AGK",
                        "Benchmark Top %",
                        "Bid Slope",
                        "Ask Slope",
                        "Spread",
                        "Cleaning Applied",
                        "Outlier Band",
                        "Rows (raw)",
                        "Rows (used)",
                        "Rows removed (outlier)",
                    ],
                    "Value": [
                        clearing_price,
                        price_method,
                        total_cost,
                        total_revenue,
                        net_cashflow,
                        price_min,
                        price_max,
                        agk,
                        int(benchmark_top_pct),
                        slope_bid,
                        slope_ask,
                        spread,
                        str(do_clean),
                        f"[{1-lower_pct:.2f}B, {1+upper_pct:.2f}B]" if do_clean else "N/A",
                        len(df_all_raw),
                        len(df_all),
                        0 if removed_df.empty else len(removed_df),
                    ],
                }
            )
            summary_df.to_excel(writer, sheet_name="Summary", index=False)

            bench_df.to_excel(writer, sheet_name="Benchmarks", index=False)
            sonuc_df.to_excel(writer, sheet_name="All_Plants", index=False)
            buyers_df.to_excel(writer, sheet_name="Buyers", index=False)
            sellers_df.to_excel(writer, sheet_name="Sellers", index=False)
            curve_df.to_excel(writer, sheet_name="Market_Curve", index=False)
            cashflow_top20.to_excel(writer, sheet_name="Cashflow_Top20", index=False)
            if not removed_df.empty:
                removed_df.to_excel(writer, sheet_name="Removed_Outliers", index=False)

            wb = writer.book

            ws_curve = wb["Market_Curve"]
            line = LineChart()
            line.title = "Market Supply–Demand Curve"
            line.y_axis.title = "Volume (tCO₂)"
            line.x_axis.title = "Price (€/tCO₂)"

            data = Reference(ws_curve, min_col=2, min_row=1, max_col=3, max_row=ws_curve.max_row)
            cats = Reference(ws_curve, min_col=1, min_row=2, max_row=ws_curve.max_row)
            line.add_data(data, titles_from_data=True)
            line.set_categories(cats)
            line.height = 12
            line.width = 24

            ws_curve["D1"] = "Carbon_Price"
            for r in range(2, ws_curve.max_row + 1):
                ws_curve[f"D{r}"] = float(clearing_price)

            line.add_data(
                Reference(ws_curve, min_col=4, min_row=1, max_row=ws_curve.max_row),
                titles_from_data=True,
            )
            ws_curve.add_chart(line, "E2")

            ws_cf = wb["Cashflow_Top20"]
            bar = BarChart()
            bar.type = "col"
            bar.title = "Top 20 Plants – ETS Net Cashflow (€)"
            bar.y_axis.title = "€"
            bar.x_axis.title = "Plant"

            data_cf = Reference(ws_cf, min_col=3, min_row=1, max_row=ws_cf.max_row)
            cats_cf = Reference(ws_cf, min_col=1, min_row=2, max_row=ws_cf.max_row)
            bar.add_data(data_cf, titles_from_data=True)
            bar.set_categories(cats_cf)
            bar.height = 12
            bar.width = 28

            bar.dataLabels = DataLabelList()
            bar.dataLabels.showVal = False

            ws_cf.add_chart(bar, "E2")

        output.seek(0)

        st.download_button(
            label="Download ETS Report (Excel + Charts)",
            data=output,
            file_name="ETS_Report_Stable_WithCharts.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # -------------------------
        # Briefing Note (Word) export
        # -------------------------
        # Optional: compute AGK=1 reference for a single example plant (e.g., Soma B) to demonstrate smoothing impact.
        soma_example = None
        try:
            sonuc_df_agk1, _, _ = ets_hesapla(
                df_all,
                price_min,
                price_max,
                1.0,  # AGK=1 reference
                slope_bid=slope_bid,
                slope_ask=slope_ask,
                spread=spread,
                benchmark_top_pct=int(benchmark_top_pct),
                price_method=price_method,
            )
            target_plant = "Soma B"
            if target_plant in set(sonuc_df_agk1.get("Plant", [])) and target_plant in set(sonuc_df.get("Plant", [])):
                cost_agk1 = float(sonuc_df_agk1.loc[sonuc_df_agk1["Plant"] == target_plant, "ets_cost_total_€"].sum())
                cost_sel = float(sonuc_df.loc[sonuc_df["Plant"] == target_plant, "ets_cost_total_€"].sum())
                soma_example = {
                    "plant": target_plant,
                    "agk1_cost_eur": f"{cost_agk1:,.0f}",
                    "agk_sel": float(agk),
                    "agk_sel_cost_eur": f"{cost_sel:,.0f}",
                }
        except Exception:
            soma_example = None

        briefing_docx = generate_briefing_note_docx(
            sonuc_df=sonuc_df,
            benchmark_map=benchmark_map,
            clearing_price=clearing_price,
            price_method=price_method,
            price_min=price_min,
            price_max=price_max,
            agk=agk,
            benchmark_top_pct=int(benchmark_top_pct),
            slope_bid=slope_bid,
            slope_ask=slope_ask,
            spread=spread,
            do_clean=do_clean,
            lower_pct=lower_pct,
            upper_pct=upper_pct,
            df_all_raw=df_all_raw,
            df_all_used=df_all,
            removed_df=removed_df,
            fx_rate=fx_rate,
            soma_example=soma_example,
        )

        st.download_button(
            label="Download Briefing Note (Word)",
            data=briefing_docx,
            file_name="ETS_Bilgi_Notu.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        csv_bytes = sonuc_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Download results as CSV",
            data=csv_bytes,
            file_name="ets_results.csv",
            mime="text/csv",
        )

    except Exception as e:
        st.error(f"Model çalışırken hata oluştu: {e}")
