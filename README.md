import os
import json
import win32com.client  # pip install pywin32

CONFIG_FILE = "config.json"

def center_text(text, width=60):
    return text.center(width)

def print_border(text=""):
    width = 60
    print("═" * width)
    if text:
        print(center_text(text, width))
        print("═" * width)

def load_last_paths():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {"input_folder": "", "output_folder": ""}

def save_last_paths(input_folder, output_folder):
    with open(CONFIG_FILE, "w") as f:
        json.dump({"input_folder": input_folder, "output_folder": output_folder}, f)

def get_directory_input(prompt_text, default_value=""):
    if default_value and os.path.exists(default_value):
        prompt = f"{prompt_text} (Enter para usar '{default_value}'): "
    else:
        prompt = f"{prompt_text}: "

    user_input = input(center_text(prompt)).strip('"')
    return user_input or default_value

def convert_excel_to_pdf(input_folder, output_folder):
    os.makedirs(output_folder, exist_ok=True)

    excel_files = [f for f in os.listdir(input_folder) if f.endswith(('.xlsx', '.xls'))]
    if not excel_files:
        print_border("⚠️ Nenhum arquivo Excel encontrado ⚠️")
        return

    print_border("📄 Iniciando Conversão com Microsoft Excel")
    print(center_text("Usando sempre a primeira aba de cada arquivo"))
    print(center_text(f"Arquivos encontrados: {len(excel_files)}"))
    print(center_text(f"Pasta de saída: {os.path.abspath(output_folder)}"))
    print("═" * 60)

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False  # Ignora prompts
    excel.EnableEvents = False   # Ignora eventos

    for index, file in enumerate(excel_files, start=1):
        input_path = os.path.join(input_folder, file)
        output_pdf = os.path.join(output_folder, os.path.splitext(file)[0] + ".pdf")

        print(center_text(f"🔄 [{index}/{len(excel_files)}] Convertendo: {file}... "), end="")

        wb = None
        try:
            wb = excel.Workbooks.Open(
                input_path,
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True
            )
            sheet = wb.Sheets(1)
            sheet.Select()
            sheet.ExportAsFixedFormat(0, output_pdf)
            print("✅")
        except Exception as e:
            print("❌")
            print(center_text(f"Erro ao converter '{file}': {e}"))
        finally:
            if wb:
                wb.Close(SaveChanges=False)

    excel.Quit()
    print_border("✅ Conversão Concluída!")

def main():
    print_border("Excel para PDF - Aba 1 (sem avisos)")
    config = load_last_paths()

    input_folder = get_directory_input("📥 Pasta com arquivos Excel", config.get("input_folder", ""))
    output_folder = get_directory_input("📤 Pasta para salvar os PDFs", config.get("output_folder", ""))

    if not os.path.exists(input_folder):
        print_border("❌ Pasta de entrada inválida!")
        return

    convert_excel_to_pdf(input_folder, output_folder)
    save_last_paths(input_folder, output_folder)

if __name__ == "__main__":
    main()
