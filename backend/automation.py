
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from io import BytesIO
import os

app = Flask(__name__)
CORS(app)

MAX_FILE_SIZE = 25 * 1024 * 1024  # 25 MB
app.config["MAX_CONTENT_LENGTH"] = MAX_FILE_SIZE


@app.route("/api/process", methods=["POST"])
def process_workbook():
    uploaded = request.files.get("file")

    if not uploaded or not uploaded.filename:
        return jsonify({"error": "Please upload an Excel file."}), 400

    if not uploaded.filename.lower().endswith(".xlsx"):
        return jsonify({"error": "Only .xlsx files are supported."}), 400

    try:
        # Load the uploaded workbook without saving over the original.
        wb = load_workbook(uploaded, data_only=False)

        if "Sheet1" not in wb.sheetnames:
            return jsonify({
                "error": "The workbook must contain a sheet named Sheet1."
            }), 400

        sheet = wb["Sheet1"]

        # Determine the actual last row containing a price in column C.
        # This supports any number of rows without hardcoding a row count.
        price_rows = [
            row for row in range(2, sheet.max_row + 1)
            if sheet.cell(row=row, column=3).value is not None
        ]

        if not price_rows:
            return jsonify({
                "error": "No prices were found in column C, starting at row 2."
            }), 400

        # Create the output column header.
        sheet.cell(row=1, column=4).value = "Discounted Price"

        processed = []

        for row in price_rows:
            original = sheet.cell(row=row, column=3).value

            # Reject text and other invalid price values.
            if isinstance(original, bool) or not isinstance(
                original, (int, float)
            ):
                return jsonify({
                    "error": f"Invalid numeric price in cell C{row}."
                }), 400

            discounted = round(original * 0.9, 2)
            sheet.cell(row=row, column=4).value = discounted

            processed.append({
                "row": row,
                "original": original,
                "discounted": discounted
            })

        # Create an Excel chart using the original and discounted prices.
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = "Original vs Discounted Prices"
        chart.y_axis.title = "Price"
        chart.x_axis.title = "Excel Row"
        chart.height = 9
        chart.width = 18

        # Use the first and last populated price rows, even when
        # the worksheet contains blank rows between them.
        first_row = min(price_rows)
        last_row = max(price_rows)

        values = Reference(
            sheet,
            min_col=3,
            max_col=4,
            min_row=1,
            max_row=last_row
        )
        chart.add_data(values, titles_from_data=True)

        # Include row labels for the categories.
        categories = Reference(
            sheet,
            min_col=1,
            min_row=first_row,
            max_row=last_row
        )
        chart.set_categories(categories)

        # Use the Excel row numbers as labels.
        for index, row in enumerate(range(first_row, last_row + 1), 1):
            sheet.cell(row=row, column=1).value = (
                sheet.cell(row=row, column=1).value or f"Row {row}"
            )

        sheet.add_chart(chart, "F2")

        # Return the workbook in memory instead of saving over the input.
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="processed_prices.xlsx",
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )

    except Exception as exc:
        app.logger.exception("Workbook processing failed")
        return jsonify({
            "error": f"Could not process the workbook: {str(exc)}"
        }), 400


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
