import os
import pandas as pd
from datetime import datetime
from utils import (
    filename_cleanup, extract_mutation_label, render_report,
    highlight_mutation_phrases, extract_red_phrase
)

def alpha_result(result: str, name: str) -> str:
    result = result.strip()

    if "(-)" in result:
        return (
            "Không phát hiện thấy các đột biến gây bệnh Alpha Thalassemia ở những vùng gen HBA đã được khảo sát."
        )

    elif "4.2" in result:
        return (
            f"Phát hiện đột biến dị hợp tử 4.2 (-α4.2/αα) trên gen HBA. {name} là người lành mang gen bệnh alpha thalassemia."
        )

    elif "3.7" in result:
        return (
            f"Phát hiện đột biến dị hợp tử 3.7 (-α3.7/αα) trên gen HBA. {name} là người lành mang gen bệnh alpha thalassemia."
        )

    elif "SEA" in result:
        return (
            f"Phát hiện đột biến dị hợp tử SEA (--SEA/αα) trên gen HBA. {name} là người lành mang gen bệnh alpha thalassemia."
        )

    else:
        return f"Nghi vấn có đột biến trên gen HBA: {alpha_result}"
    
def beta_result(result: str, name: str) -> str:
    result = result.strip().lower()

    if result == "bt":
        return "Không phát hiện thấy các đột biến gây bệnh Beta Thalassemia ở những vùng gen HBB đã được khảo sát."

    elif "dị hợp" in result:
        mutation = result.replace("dị hợp", "").strip().upper()
        return f"Phát hiện đột biến dị hợp tử {mutation} trên gen HBB. {name} là người mang gen bệnh beta thalassemia."

    else:
        return f"Nghi vấn có đột biến trên gen HBB: {result}"
    
def process_thalassemia_excel(file_path, output_dir):
    from datetime import datetime
    df = pd.read_excel(file_path, header=None)
    results = []
    today = datetime.today()
    for _, row in df.iterrows():
        context = {
            "ID": row[2],
            "name": row[3],
            "yob": row[4],
            "alpha_mutation_result": alpha_result(row[18], row[3]),
            "beta_mutation_result": beta_result(row[19], row[3]),
            "date": str(today.day),
            "month": str(today.month),
            "year": str(today.year)
        }
        alpha_mut = extract_mutation_label(context["alpha_mutation_result"])
        beta_mut = extract_mutation_label(context["beta_mutation_result"])
        mutation_part = "_".join(filter(None, [alpha_mut, beta_mut]))

        id_clean = filename_cleanup(str(row[2]))
        name_clean = filename_cleanup(str(row[3]))
        if mutation_part:
            output_name = f"{id_clean}_{name_clean}_{mutation_part}_Thalassemia"
        else:
            output_name = f"{id_clean}_{name_clean}_Thalassemia"
        output_name = filename_cleanup(output_name)

        output_path = render_report("Thalassemia", context, output_name, output_dir)
        phrases = [
            extract_red_phrase(context["alpha_mutation_result"]),
            extract_red_phrase(context["beta_mutation_result"])
        ]
        phrases = [p for p in phrases if p]
        highlight_mutation_phrases(output_path, phrases)
        results.append((output_name, output_path))

    return results

if __name__ == "__main__":
    test_excel = "thal.xlsx"
    output_dir = "test_output"
    os.makedirs(output_dir, exist_ok=True)

    results = process_thalassemia_excel(test_excel, output_dir)

    for output_name, output_path in results:
        print(f"Đã xuất file {output_name}:")