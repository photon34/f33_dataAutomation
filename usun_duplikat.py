import pandas as pd
import os
import re

input = 'Revil Leads.xlsx'
output = 'firmy.xlsx'

def cleanMoney(money):
    if pd.isna(money):
        return 0
    if isinstance(money, str):
        money = money.strip().upper().replace(" ", "")
        money = re.sub(r'[^0-9.BMK]', '', money)
        try:
            if 'B' in money:
                return float(money.replace('B', '')) * 1_000_000_000
            elif 'M' in money:
                return float(money.replace('M', '')) * 1_000_000
            elif 'K' in money:
                return float(money.replace('K', '')) * 1_000
            else:
                return float(money)
        except:
            return 0
    return money

def createOutput():
    try:
        allInput = pd.read_excel(input, sheet_name=None)
        res = []
        for name, df in allInput.items():
            if all(k in df.columns for k in ['Company', 'Recent Funding Amount', 'Using cloud marketplaces?']):
                filtered = df[['Company', 'Recent Funding Amount', 'Using cloud marketplaces?']].drop_duplicates()
                res.append(filtered)
            else:
                print(f"Pominięto arkusz '{name}' – brak wymaganych kolumn.")


        joint = pd.concat(res, ignore_index=True)
        nonReps = joint.drop_duplicates(subset='Company').copy()
        nonReps['Recent Funding Amount (USD)'] = nonReps['Recent Funding Amount'].apply(cleanMoney)
        nonReps = nonReps.sort_values(by='Recent Funding Amount (USD)', ascending=False)
        nonReps = nonReps.drop(columns=['Recent Funding Amount'])

        nonReps.to_excel(output, index=False)
        print(f"Zapisano {len(nonReps)} firm do pliku: {output}")
        print(f"Folder: {os.getcwd()}")

    except FileNotFoundError:
        print(f"Nie znaleziono pliku '{input}'")
    except Exception as e:
        print(f"Błąd: {e}")

if __name__ == '__main__':
    createOutput()
