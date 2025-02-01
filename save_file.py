import asyncio
import difflib
import pprint
from rebrowser_playwright.async_api import async_playwright
import time
import pandas as pd
import os


CHROME_ARGS = ['--disable-blink-features=AutomationControlled']
arquivos = os.listdir('./anexos')

#ler arquivo xlsx de forma async
async def read_xlsx(file_path):
    df = pd.read_excel(file_path)
    return df

#buscar arquivo com nome da empresa na pasta de anexos
def find_file(list_empresas):
    for empresa in list_empresas:
        if empresa.name.endswith(".txt"):
            return empresa
        

async def find_closest_match(name, arquivos):
    arquivo = await asyncio.to_thread(difflib.get_close_matches, name, arquivos, n=1, cutoff=0.5)
    return arquivo[0] if arquivo else None
        
async def config_list_empresas(list_empresas, arquivos):
    result = []
    tasks = [find_closest_match(name, arquivos) for name in list_empresas]

    arquivos_correspondentes = await asyncio.gather(*tasks)

    for name, arquivo in zip(list_empresas, arquivos_correspondentes):
        if arquivo:
            item = {
                "name": name,
                "file": "./anexos/" + arquivo
            }
            result.append(item)

    return result

async def register_odoo(list_empresas):
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, channel="chrome", devtools=False, slow_mo=200, args=CHROME_ARGS)

        page = await browser.new_page()
        await page.set_viewport_size({"width": 1080, "height": 720})
        await page.goto("https://tributojusto.odoo.com/web/login")

        await page.locator('css=#login').type("suporte@tributojusto.com.br")
        await page.locator('css=#password').type("8j^Xkyovzz&PN^")
        await page.locator('css=button.btn.btn-primary[type="submit"]').click()
        # time.sleep(5)
        await page.goto('https://tributojusto.odoo.com/odoo/helpdesk/1/action-984')

        for empresa in list_empresas[77:]:
            print(empresa)
            input_search = page.locator('//html/body/div[1]/div/div[1]/div/div[2]/div/div/div/input')
            await input_search.press('Backspace')
            await input_search.type(empresa['name'])
            await page.locator('css=.d-print-none.btn.border-0.p-0[role="button"]').click()
            time.sleep(1)

            await page.locator("css=strong.o_kanban_record_title > span", has_text=empresa['name']).click()
            time.sleep(1)
            
            # button = page.locator("//html/body/div[1]/div/div/div[2]/div/div[2]/div/div[1]/div/div/span[2]/button")
            # await page.wait_for_selector("//html/body/div[1]/div/div/div[2]/div/div[2]/div/div[1]/div/div/span[2]/button", timeout=60000)
            # await button.scroll_into_view_if_needed()
            # if await button.is_visible():
            #      await button.click(force=True)
            # else:
            #      print("O botão não está visível ou não está clicável")
            input_file = page.locator("//html/body/div[1]/div/div/div[2]/div/div[2]/div/div[1]/div/div/input")
             # await page.evaluate('''(input, filePath) => {
             #     input.style.display = "block";  // Mostra o campo se necessário
             #     input.files = new DataTransfer();
             #     input.files.items.add(new File([filePath], "nome_do_arquivo"));
             # }''', input_file, empresa['file'])
            await input_file.set_input_files(empresa['file'])
            # await input_search.press('Esc')
            time.sleep(1)
            await page.locator("//html/body/div[1]/div/div/div[1]/div/div[1]/div[2]/ol/li[3]/a").click()
            

            time.sleep(2)
        await browser.close()
    



async def init():
    data = await read_xlsx("IMPORTAÇÃO - ODOO - TESTE JHOW (20).xlsx")
    list_companys = await config_list_empresas(data['name'], arquivos)
    pprint.pprint(list_companys[89:][0])
    # await register_odoo(list_companys)

asyncio.run(init())
