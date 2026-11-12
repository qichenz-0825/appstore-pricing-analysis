"""Build the interactive FP&A dashboard (reports/dashboard.html) from reports/model_inputs.json."""
from datetime import date
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
REPORTS = ROOT / "reports"


def main():
    inputs = (REPORTS / "model_inputs.json").read_text()
    n_apps = pd.read_csv(ROOT / "data" / "apps.csv").trackId.nunique()
    html = (Path(__file__).parent / "dashboard_template.html").read_text()
    html = (html.replace("{{INPUTS}}", inputs)
                .replace("{{N_APPS}}", f"{n_apps:,}")
                .replace("{{DATE}}", date.today().isoformat()))
    (REPORTS / "dashboard.html").write_text(html)
    print(f"Wrote {REPORTS / 'dashboard.html'}")


if __name__ == "__main__":
    main()
