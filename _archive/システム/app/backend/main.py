import csv

file_path = "/Users/youichi/Desktop/配送AI/顧客データ/data.csv"

with open(file_path, encoding="utf-8") as f:
    reader = csv.DictReader(f)

    print("=== 顧客一覧 ===")

    for row in reader:
        print(f"{row['顧客コード']} : {row['名前']}")


