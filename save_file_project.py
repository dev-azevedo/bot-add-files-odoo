import asyncio
import difflib
import pprint
from rebrowser_playwright.async_api import async_playwright
import time
import pandas as pd
import os


CHROME_ARGS = ['--disable-blink-features=AutomationControlled']
arquivos = os.listdir('./inss-andamento')

"""
17 - Ferrari
18 - Compensação
19 - McLaren
20 - Mercedes-Benz
"""
list_project_id = [18, 17, 19, 20]

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
        for arquivo in arquivos_correspondentes:
            caminho = f"./inss-andamento/{arquivo}"
            
            try:
                arquivos = os.listdir(caminho)
                arquivos_encontrados = [os.path.join(caminho, f)  for f in arquivos]

                if arquivos_encontrados:
                    result.append({"name": name, "files": arquivos_encontrados})
            except:
                continue
          
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
        
        
        await page.wait_for_url("https://tributojusto.odoo.com/odoo", timeout=60000)
        # time.sleep(5)
        
        for project_id in list_project_id:
            await page.goto(f'https://tributojusto.odoo.com/odoo/project/{project_id}/tasks')

            for empresa in list_empresas:
                print(empresa)
                input_search = page.locator('//html/body/div[1]/div/div[1]/div/div[2]/div/div/div/input')
                await input_search.press('Backspace')
                await input_search.type(empresa['name'])
                await page.locator('css=.d-print-none.btn.border-0.p-0[role="button"]').click()
                time.sleep(1)
                
               
                # card_company = page.locator("css= body>div.o_action_manager>div>div.o_content>div>div>article>div.oe_kanban_color_0.oe_kanban_card.oe_kanban_global_click>div.oe_kanban_content > div.o_kanban_record_body.text-muted > span:nth-child(2)", has_text=empresa['name'])

                try:
                    await page.locator("css=div.o_kanban_record_body > span.o_text_block", has_text=empresa['name']).click(timeout=5000)
                    print("Card encontrado ✅")
                    time.sleep(1)
                
                    input_file = page.locator("//html/body/div[1]/div/div/div[2]/div/div[2]/div/div[1]/div/div/input") 

                    arquivos_validos = [file for file in empresa['files'] if os.path.isfile(file)]
                    
                    if arquivos_validos:
                        await input_file.set_input_files(arquivos_validos)
                        time.sleep(1)
                        
                    time.sleep(1)
                    await page.locator("//html/body/div[1]/div/div/div[1]/div/div[1]/div[2]/ol/li[3]/a").click()
                

                    time.sleep(2)
        
                except Exception as e:
                    print(e)
                    continue
                    
            await browser.close()
    



async def init():
    data = await read_xlsx("inss-andamento.xlsx")
    print(data['name'])
    list_companys = await config_list_empresas(data['name'], arquivos)
    pprint.pprint(list_companys)
    await register_odoo(list_companys)

asyncio.run(init())
