import os
from xml.sax.saxutils import escape

# =========================
# CONFIGURATION
# =========================
OUTPUT_DIR = "out_trades"          # Output directory
FILES_TO_GENERATE = 5              # 3 files per run
TRADES_PER_FILE = 100              # 100 entries per txt
VALUE_DATE = "20260309"            # Fixed value date as requested

# Three currency pairs (rotates USDJPY -> USDALL -> USDINR -> repeat)
# Format: (quote_pair, purchase_ccy, sale_ccy, fixed_rate)
PAIRS = [
    ("USDJPY", "JPY", "USD", 115.79795),
    ("USDALL", "ALL", "USD", 81.61578),
    ("USDINR", "INR", "USD", 80.53314),
]

# Amount ranges (purchase is non-USD; sale is USD)
PURCHASE_AMOUNT_MIN = 200_000.0
PURCHASE_AMOUNT_MAX = 10_000_000.0

# Static fields
CUSTOMER_NUMBER = "BOA"
TRADE_TYPE = "S"
INTERFACE_ID = "FXTRADEMIG"
AREA = "EXOTIC"
PORTFOLIO = "3PTY"
MESSAGE_TYPE = "FX"
STANDARD_SETTLEMENT = "Y"
TRADER = "WSS"

# External deal number (EDN) uniqueness:
# We'll persist a counter at OUTPUT_DIR/edn_counter.txt and transform it into:
# 4 uppercase letters + 7 digits + '/' + 1 digit, e.g., ABCD1234567/4
EDN_COUNTER_FILE = "edn_counter.txt"
EDN_START = 1_000_000  # starting seed if no file exists


# =========================
# HELPERS
# =========================
def ensure_dir(path):
    os.makedirs(path, exist_ok=True)

def read_counter(path, default_value):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return int(f.read().strip())
    except Exception:
        return default_value

def write_counter(path, value):
    with open(path, "w", encoding="utf-8") as f:
        f.write(str(value))

def counter_to_letters(n, width=4):
    """
    Convert integer n to a base-26 'letters' string (A-Z), fixed width.
    Deterministic mapping to make EDNs look random-like but guaranteed unique.
    """
    letters = []
    for _ in range(width):
        n, rem = divmod(n, 26)
        letters.append(chr(ord('A') + rem))
    # We built little-endian; reverse to get a 'normal' look
    return "".join(reversed(letters))

def make_external_deal_number(counter):
    """
    Create unique EDN with the desired format:
    4 letters + 7 digits + '/' + digit (1..9)
    """
    prefix = counter_to_letters(counter, width=4)
    digits = f"{counter % 10_000_000:07d}"
    suffix = (counter % 9) + 1  # 1..9
    return f"{prefix}{digits}/{suffix}"

def fmt_amount(x):
    return f"{x:.5f}"

def make_fxtrade_xml(value_date, quote_pair, purchase_ccy, sale_ccy, trade_rate,
                     purchase_amount, sale_amount, external_deal_number):
    # cname = <customer_number><portfolio><value_date><external_deal_number>
    cname = f"{CUSTOMER_NUMBER}{PORTFOLIO}{value_date}{external_deal_number}"
    E = lambda s: escape(str(s))
    xml = []
    xml.append("  <FXTRADE>")
    xml.append(f"    <customer_number>{E(CUSTOMER_NUMBER)}</customer_number>")
    xml.append(f"    <trade_type>{E(TRADE_TYPE)}</trade_type>")
    xml.append(f"    <quote_pair>{E(quote_pair)}</quote_pair>")
    xml.append(f"    <purchase_ccy>{E(purchase_ccy)}</purchase_ccy>")
    xml.append(f"    <purchase_amount>{E(fmt_amount(purchase_amount))}</purchase_amount>")
    xml.append(f"    <sale_ccy>{E(sale_ccy)}</sale_ccy>")
    xml.append(f"    <sale_amount>{E(fmt_amount(sale_amount))}</sale_amount>")
    xml.append(f"    <trade_rate>{E(f'{trade_rate:.5f}')}</trade_rate>")
    xml.append(f"    <value_date>{E(value_date)}</value_date>")
    xml.append(f"    <external_deal_number>{E(external_deal_number)}</external_deal_number>")
    xml.append(f"    <interface_id>{E(INTERFACE_ID)}</interface_id>")
    xml.append(f"    <area>{E(AREA)}</area>")
    xml.append(f"    <portfolio>{E(PORTFOLIO)}</portfolio>")
    xml.append(f"    <message_type>{E(MESSAGE_TYPE)}</message_type>")
    xml.append(f"    <standard_settlement>{E(STANDARD_SETTLEMENT)}</standard_settlement>")
    xml.append(f"    <cname>{E(cname)}</cname>")
    xml.append(f"    <trader>{E(TRADER)}</trader>")
    xml.append("  </FXTRADE>")
    return "\n".join(xml)

def gen_amounts_for_pair(rate):
    """
    Generate purchase_amount (non-USD) and sale_amount (USD).
    purchase ~ [min, max]; sale = purchase / rate +/- ~1% noise; both rounded to integers.
    """
    import random
    purchase_amt = random.uniform(PURCHASE_AMOUNT_MIN, PURCHASE_AMOUNT_MAX)
    sale_amt = purchase_amt / rate
    sale_amt *= random.uniform(0.99, 1.01)
    purchase_amt = round(purchase_amt)
    sale_amt = round(sale_amt)
    return float(purchase_amt), float(sale_amt)


# =========================
# MAIN
# =========================
def main():
    ensure_dir(OUTPUT_DIR)
    counter_path = os.path.join(OUTPUT_DIR, EDN_COUNTER_FILE)
    counter = read_counter(counter_path, EDN_START)

    for file_idx in range(1, FILES_TO_GENERATE + 1):
        lines = []
        lines.append('<?xml version="1.0" encoding="UTF-8"?>')
        lines.append("<wss-request>")

        for t in range(TRADES_PER_FILE):
            pair = PAIRS[t % len(PAIRS)]
            quote_pair, purchase_ccy, sale_ccy, rate = pair

            # Unique EDN based on persistent counter
            edn = make_external_deal_number(counter)
            counter += 1

            # Amounts for non-USD purchase, USD sale
            purchase_amt, sale_amt = gen_amounts_for_pair(rate)

            fx_xml = make_fxtrade_xml(
                value_date=VALUE_DATE,
                quote_pair=quote_pair,
                purchase_ccy=purchase_ccy,
                sale_ccy=sale_ccy,
                trade_rate=rate,
                purchase_amount=purchase_amt,
                sale_amount=sale_amt,
                external_deal_number=edn
            )
            lines.append(fx_xml)

        lines.append("</wss-request>")
        content = "\n".join(lines)

        out_path = os.path.join(OUTPUT_DIR, f"fx_trades_{file_idx:03d}.txt")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(content)

    # Persist updated counter (guarantees no repeats on rerun)
    write_counter(counter_path, counter)

    print(f"Done. Generated {FILES_TO_GENERATE} file(s) in '{OUTPUT_DIR}' with {TRADES_PER_FILE} FXTRADE entries each.")
    print(f"Value date fixed to {VALUE_DATE}. Three pairs rotated: {', '.join(p[0] for p in PAIRS)}.")
    print(f"External deal numbers are unique across reruns (counter persisted at {counter_path}).")

if __name__ == "__main__":
    main()