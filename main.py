import json
import time
import requests
from bs4 import BeautifulSoup
import pandas as pd
from chat import make_chat


def name_process(name: str) -> list:
    if not name or name.strip() == "":
        return ["", "", ""]
    name = name.split(" ")
    match len(name):
        case 0:
            fn = ""
            mn = ""
            ln = ""
        case 1:
            fn = name[0]
            mn = ""
            ln = ""
        case 2:
            fn = name[0]
            mn = ""
            ln = name[1]
        case 3:
            fn = name[0]
            mn = name[1]
            ln = name[2]
        case _:
            fn = name[0]
            mn = name[1]
            ln = ""
            for i in range(2, len(name)):
                ln += name[i] + " "
            ln = ln[:-1]
    return [fn, mn, ln]


if __name__ == "__main__":
    try:
        with open("data.json", "r") as f:
            data_list = json.load(f)
    except FileNotFoundError:
        data_list = []

    url = "https://sheet2site.com/api/v3/index.php?key=1QjkZQqZiBPNOiEgsR0PlARV275MgdYt0qY00x_cv1fs&g=1&e=1&e=1"
    response = requests.get(url)
    text = response.text
    soup = BeautifulSoup(text, "html.parser")
    table_list = soup.find("table", {"id": "example"}).find("tbody").find_all("tr")
    for table in table_list:
        td = table.find_all("td")
        company_name = td[1].text
        website = td[6].text
        ceo_name = td[-1].text
        try:
            ceo_name_info = name_process(ceo_name)
            first_name = ceo_name_info[0]
            middle_name = ceo_name_info[1]
            last_name = ceo_name_info[2]
        except IndexError as e:
            # Rarely seen error
            print(f"CEO NAME ERROR: {e}; Company Name: {company_name}; CEO Name: {ceo_name}")
            first_name = None
            middle_name = None
            last_name = None
            time.sleep(5)
        data = {
            "Company": company_name,
            "Website": f"https://{website}" if not website.startswith("http") else website,
            "CEO First Name": first_name,
            "CEO Middle Name": middle_name,
            "CEO Last Name": last_name,
        }

        cmo_task_not_done = True
        for company_dict in data_list:
            if company_dict["Company"] == company_name:
                cmo_task_not_done = False
        if cmo_task_not_done:
            try:
                # Do not use f-string here
                prompt = 'I am searching for CMO information of Fortune 500, I give you company name. You return a Json to me, the Json format is shown below: \n```\n{\n   \"companyName\": \"\",\n   \"CMOName\": \"\",\n   \"CMOEmail\": \"\"\n}\n\nDO NOT INCLUDE ANY OTHER CONTENT IN YOUR RESPONSE OTHER THAN THE JSON. Use None IF INFO IS NOT AVAILABLE.\n\nThe company name is company_name.'
                prompt = prompt.replace("company_name", company_name)
                cmo_data = json.loads(make_chat(prompt))
                cmo_name = name_process(cmo_data["CMOName"])
                cmo_email = cmo_data["CMOEmail"]
                data["CMO First Name"] = cmo_name[0]
                data["CMO Middle Name"] = cmo_name[1]
                data["CMO Last Name"] = cmo_name[2]
                data["CMO Email"] = cmo_email
            except (json.decoder.JSONDecodeError, RuntimeError) as e:
                print(f"CMO Error: {e}")
                data["CMO First Name"] = None
                data["CMO Middle Name"] = None
                data["CMO Last Name"] = None
                data["CMO Email"] = None
                time.sleep(3)
            data_list.append(data)
            time.sleep(10)
        else:
            print(f"Company {company_name} already in list, skip CMO task.")
        with open("data.json", "w+") as f:
            json.dump(data_list, f, indent=4)
    df = pd.DataFrame(data_list, index=None)
    df.to_excel("data.xlsx")
