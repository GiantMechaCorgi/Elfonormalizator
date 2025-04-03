import sys
import pandas
import os
import colorama
from colorama import Fore, Style

def input_name():
    name = input(f"Wprowadz nazwę próbki odniesienia dla {file}: ")
    for lista in elf_list:
        if name == lista[0]:
            print(f"Próbka odniesienia: {name}.")
            return name
    print(Fore.YELLOW + "Nieprawidłowa nazwa próbki.\n" + Style.RESET_ALL)
    return input_name()

colorama.init()

print(Fore.GREEN + "################################################################\n"
      " Pamiętaj, żeby pliki w folderze data miały rozszerzenie .xlsx!\n"
      "################################################################\n\n"  + Style.RESET_ALL)

if not os.path.exists("./output"):
    print(f"Tworzenie folderu output")
    os.makedirs("./output")
if not os.path.exists("./data"):
    print(Fore.YELLOW + f"Tworzenie folderu data. Umieść w nim swoje pliki z rozszerzeniem .xlsx i uruchom program ponownie." + Style.RESET_ALL)
    os.makedirs("./data")
    input("Enter by zamknąć.")
    sys.exit()

print("Lista odnalezionych plików: ")
for file in os.listdir("./data"):
    print(file)
print("\n")

for file in os.listdir("./data"):
    if file.endswith(".xlsx"):
        print(f"Obliczenia dla pliku {file}.")
        # ========== LOADING RAW DATA FROM FILE ==========
        try:
            data = pandas.read_excel(f"./data/{file}", engine="openpyxl", header=7)
        except FileNotFoundError:
            print(Fore.RED + f"Błąd: Plik {file} nie istnieje i zostanie pominięty." + Style.RESET_ALL)
            continue
        except ValueError as e:
            print(Fore.RED + f"Błąd wczytywania {file}: {e}. \nPlik zostanie pominięty." + Style.RESET_ALL)
            continue
        except Exception as e:
            print(Fore.RED + f"Niespodziewany błąd podczas wczytywania pliku {file}: {e}. \nPlik zostanie pominięty." + Style.RESET_ALL)
            continue

        # Usuwanie zawartości oddzielonej od tabeli pustym wierszem
        for i, row in data.iterrows():
            if row.isnull().all():
                data = data.iloc[:i]

        # Sprawdzenie zawartości pliku
        if data is None or data.empty:
            print(Fore.RED + f"Ostrzeżenie: {file} Jest pusty lub niepoprawnie sformatowany i zostanie pominięty." + Style.RESET_ALL)
            continue

        expected_columns = ["Sample Name", "Target Name", "Quantity"]
        missing_columns = [col for col in expected_columns if col not in data.columns]
        if missing_columns:
            print(Fore.RED + f"Błąd: {file} nie posiada spodziewanych kolumn: {missing_columns}. \nSprawdź, czy dane są odpowiednio sformatowane. Plik zostanie pominięty." + Style.RESET_ALL)
            continue


        wells = []
        for index, row in data.iterrows():
            if not pandas.isna(row.to_dict()["Sample Name"]):
                wells.append([row.to_dict()["Sample Name"], row.to_dict()["Target Name"], row.to_dict()["Quantity"]])
        #print(wells)

        # ========== CALCULATIONS ==========

        # # ========== LICZENIE ELF ==========

        # tam gdzie target = elf, zrobic srednia z następujących po sobie elfów

        wells_sorted = sorted(wells, key=lambda x: (x[0], x[1]))
        # print(wells_sorted)

        # lista list [nazwa próbki, średnia z jej elfów]
        elf_list = []
        well_name = wells_sorted[0][0]
        suma = 0
        count = 0
        for well in wells_sorted:
            if well[0] != well_name:
                try:
                    elf_list.append([well_name, suma / count])
                except ZeroDivisionError:
                    print(Fore.RED + f"Ostrzeżenie: Nie znaleziono targetów elf dla {well_name}. Średnią wartość ustawiono na NaN." + Style.RESET_ALL)
                    elf_list.append([well_name, float('nan')])
                well_name = well[0]
                count = 0
                suma = 0
            if well[1] == 'elf':
                count += 1
                suma += well[2]
            if well == wells_sorted[-1]:
                try:
                    elf_list.append([well_name, suma / count])
                except ZeroDivisionError:
                    print(Fore.RED + f"Ostrzeżenie: Nie znaleziono targetów elf dla {well_name}. Średnią wartość ustawiono na NaN.\nSprawdź, czy plik z danymi jest poprawny." + Style.RESET_ALL)
                    elf_list.append([well_name, float('nan')])

        # print(elf_list)

        # ========== NORMALIZACJA DO ELF ==========

        # uzupelniamy listę wells_sorted wartosciami nieelfowymi podzielonymi przez sredniego elfa tej probki
        for well in wells_sorted:
            if well[1] != 'elf':
                elf = [item[1] for item in elf_list if item[0] == well[0]]
                value = well[2] / elf[0]
                well.append(value)

        # print(wells_sorted)
        # ========== OBLICZENIE PRÓBKI ODNIESIENIA ==========

        sample_name = input_name()

        suma = 0
        count = 0
        target_name = ""
        normal_targets = []
        avg = 0
        for well in wells_sorted:
            if well[0] == sample_name and well[1] != 'elf':
                target_name = well[1]
                break

        for well in wells_sorted:
            if well[0] == sample_name and well[1] != 'elf':
                if well[1] != target_name:
                    if count > 0:
                        avg = suma / count
                    else:
                        avg = 0
                    normal_targets.append([sample_name, target_name, avg])
                    target_name = well[1]
                    suma = 0
                    count = 0

                suma += well[3]
                count += 1

        avg = suma / count
        normal_targets.append([sample_name, target_name, avg])

        # print(normal_targets)



        # ========== NORMALIZACJA DO PRÓBKI ODNIESIENIA ==========

        list1 = [[well[0], well[1], well[3] / next((nt[2] for nt in normal_targets if nt[1] == well[1]), 1)] for well in wells_sorted if (well[0] != sample_name and well[1] != 'elf')]
        # print(f"list1: {list1}")



        # ========== OBLICZANIE ŚREDNICH Z POWTÓRZEŃ TECHNICZNYCH ==========

        grouped = {name: [[j, value] for n, j, value in list1 if n==name] for name, i, j in list1}
        # print(f"grouped: {grouped}")

        grouped_averages = {
            key: [
                [target, sum(v for t, v in values if t == target) / len([v for t, v in values if t == target])]
                for target in {t for t, v in values}
            ]
            for key, values in grouped.items()
        }

        # print(f"grouped_averages: {grouped_averages}")

        # niepotrzebne bo w dict nie ma probek z sample_name
        # grouped_averages = {key:[[target, 1.0 if key == sample_name else value] for target, value in values] for key, values in grouped_averages.items()}

        # Doklejenie na koniec grouped_averages nieelfowych targetow probki odniesienia
        list_probka_odniesienia = [[well[0], well[1], well[3] / next((nt[2] for nt in normal_targets if nt[1] == well[1]), 1)] for well in wells_sorted if (well[0] == sample_name and well[1] != 'elf')]
        targety_probki_odniesienia = set(item[1] for item in list_probka_odniesienia)
        # print(targety_probki_odniesienia)
        grouped_averages[sample_name] = []
        for target in targety_probki_odniesienia:
            grouped_averages[sample_name].append([target, 1.0])

        # Stara implementacja, nie działa jeśli pierwszy well nie ma wszystkich targetów
        # first_target = next(iter(grouped_averages))
        # grouped_averages[sample_name] = []
        # for target in grouped_averages[first_target]:
        #     # print(target[0])
        #     grouped_averages[sample_name].append([target[0], 1.0])

        # print(grouped_averages)
        # ========== ZAPISYWANIE DO PLIKU ==========
        output_file = f"./output/{file.split('.')[0]}_obliczenia.xlsx"

        dataframes = [] # jednak za duzo roboty z formatowaniem ale niech zostanie
        for key, values in grouped_averages.items():
            # Creating a dataframe for each sample name
            dataframe = pandas.DataFrame(values, columns=['Target Name', 'Average'])
            dataframe['Sample Name'] = key
            dataframe = dataframe[["Sample Name", "Target Name", "Average"]] # reordering the columns
            dataframes.append(dataframe)

        combined_dataframe = pandas.concat(dataframes, ignore_index=True)

        wells_sorted_df = pandas.DataFrame(wells_sorted, columns=['Sample Name', 'Target', 'Quantity', 'Quant / Elf avg'])

        elf_list_df = pandas.DataFrame(elf_list, columns=['Sample', 'Elf Avg Value'])

        normal_targets_df = pandas.DataFrame(normal_targets, columns=['Sample Name', 'Target', 'Normalized Value'])

        try:
            with pandas.ExcelWriter(output_file, engine="xlsxwriter") as writer:
                combined_dataframe.to_excel(writer, sheet_name="Results", index=False)
                wells_sorted_df.to_excel(writer, sheet_name="Results", startrow=len(combined_dataframe) + 3, startcol=0, index=False)
                elf_list_df.to_excel(writer, sheet_name="Results", startrow=len(combined_dataframe) + len(wells_sorted_df) + 6, startcol=0, index=False)
                normal_targets_df.to_excel(writer, sheet_name="Results", startrow=len(combined_dataframe) + len(wells_sorted_df) + len(elf_list_df) + 9, startcol=0, index=False)

                # Formatowanie spreadsheeta dla czytelności
                workbook = writer.book
                worksheet = writer.sheets["Results"]
                center_format = workbook.add_format({"align": "center", "valign": "vcenter"})

                worksheet.set_column('A:A', 15)
                worksheet.set_column('B:B', 12, center_format)
                worksheet.set_column('C:C', 10, workbook.add_format({'num_format': '0.0000', "align": "right"}))
                worksheet.set_column('D:D', 15, workbook.add_format({'num_format': '0.0000', "align": "right"}))
        except PermissionError:
            print(Fore.RED + f"Błąd: Nie można zapisać do pliku {output_file}. Możliwe, że plik jest otwarty. Zamknij plik i spróbuj ponownie." + Style.RESET_ALL)
            continue
        except Exception as e:
            print(Fore.RED + f"Nieoczekiwany błąd przy zapisie do pliku {output_file}: {e}" + Style.RESET_ALL)
            continue

        print(f"Zapisano wyniki dla {file} w {output_file}.\n\n")

input(Fore.GREEN + "-----------------------------------------------\nProgram zakończył pracę, sprawdź folder output.\nEnter by zamknąć.\n-----------------------------------------------" + Style.RESET_ALL)








