from pathlib import Path

from flask import Flask, jsonify, render_template, request

from data_engine import INTERVALS, SPXAnalyzer


app = Flask(__name__)
analyzer = SPXAnalyzer(Path(__file__).with_name("sp500.xlsx"))


def _interval_arg() -> str:
    interval = request.args.get("interval", "daily").lower()
    if interval not in INTERVALS:
        interval = "daily"
    return interval


@app.route("/")
def index():
    return render_template("index.html", intervals=INTERVALS, freshness=analyzer.data_freshness())


@app.route("/api/series")
def series():
    return jsonify(analyzer.get_series(_interval_arg()))


@app.route("/api/dashboard")
def dashboard():
    return jsonify(analyzer.get_dashboard(_interval_arg()))


@app.route("/api/events")
def events():
    return jsonify(analyzer.get_events(_interval_arg()))


@app.route("/api/event/<event_id>")
def event_detail(event_id: str):
    try:
        return jsonify(analyzer.get_event_detail(_interval_arg(), event_id))
    except KeyError:
        return jsonify({"error": "event_not_found"}), 404


@app.route("/api/summary")
def summary():
    return jsonify(analyzer.get_summary(_interval_arg()))


if __name__ == "__main__":
    app.run(debug=True)
