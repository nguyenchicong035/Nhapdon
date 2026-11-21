import re
from decimal import Decimal
from pathlib import Path
from parsers.common import norm, quantity


def apply(items, path, config, default, auto):
    store = default
    if auto:
        filename = " " + norm(Path(path).stem) + " "
        matches = [s for s, wh in config["warehouses"].items()
                   if any(" " + norm(k) + " " in filename for k in wh["keywords"])]
        if len(matches) == 1:
            store = matches[0]
    code = config["warehouses"][store]["code"]
    adapter = config["adapter"]
    result = []
    for original in items:
        row = dict(original, warehouse=code, source=Path(path).name)
        names = [norm(row.get(k)) for k in ("name", "variant")]
        match = next((m for name in names if (m := re.search(r"\bcombo\s*0*(\d+)\b", name))), None)
        count = quantity(row["qty"])
        sku = row.get("sku", "")
        if row.get("price"):
            row["price"] = Decimal(str(row["price"]).replace(",", "."))
        override = config.get("sku_multipliers", {}).get(sku)
        if match or override:
            factor = quantity(override or match[1])
            count *= factor
            if config["combo_sku_mode"] == "map":
                if sku not in config["combo_to_unit"]:
                    raise ValueError(f"Thiếu SKU lẻ cho combo {sku!r}.")
                sku = config["combo_to_unit"][sku]
            # The platform unit price describes a combo. Preserve the line total.
            if row.get("price"):
                price = Decimal(str(row["price"]))
                row["amount"] = price * quantity(original["qty"])
                row["price"] = price / factor
        row["_source_sku"] = original.get("sku", "")
        row["_mapping_key"] = sku
        row.update(sku=config["sku_to_misa"].get(sku, sku), qty=count)
        entry = config.get("catalog", {}).get(row["sku"])
        if entry:
            row.update(name=entry["name"], unit=entry["unit"])
        if sku in config.get("review_skus", []):
            row["sku"] = ""
            row["name"] = f"[{sku}] {row.get('name', '')}"
        row["_adapter_units"] = count if adapter["enabled"] and any(
            norm(k) in name for k in adapter["keywords"] if norm(k) for name in names) else 0
        result.append(row)
    return result


def accessories(rows, cfg):
    grouped = {}
    for row in rows:
        if row.get("_adapter_units") and row.get("order"):
            grouped.setdefault(row["order"], []).append(row)
    for order, group in grouped.items():
        base = group[0]
        n = sum(r["_adapter_units"] for r in group) if cfg["mode"] == "per_unit" else 1
        details = {k: base.get(k, "") for k in ("date", "customer", "buyer")}
        yield dict(**details, order=order, sku=cfg["sku"], name=cfg["name"], unit=cfg["unit"],
                   warehouse=base["warehouse"], qty=n*quantity(cfg["quantity"]), source="Adapto tự thêm")
