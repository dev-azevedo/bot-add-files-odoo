import asyncio
import difflib
import pprint
from rebrowser_playwright.async_api import async_playwright
import time
import pandas as pd
import os


CHROME_ARGS = ['--disable-blink-features=AutomationControlled']
company_igonre = ['POSTO BOM RETIRO LTDA - 80.376.478/0001-50',
'SUDASEG BENEFICIOS E PROMOCOES DE VENDAS LTDA - 16.538.051/0001-01',
'BELEM RIO SEGURANCA LTDA - 17.433.496/0001-90',
'NELOR CHURRASCARIA E RESTAURANTE LTDA - 02.645.436/0001-63',
'TRANSPORTES LEISI LTDA - 95.832.150/0001-78',
'AUTO POSTO OESTE VERDE LTDA - 80.359.003/0001-55']

async def read_xlsx(file_path):
    df = pd.read_excel(file_path)
    df = pd.DataFrame(df)
    return df
        
        
async def config_list_empresas(list_empresas):
    result = []

    for line in list_empresas.itertuples():
        followres = [name.strip() for name in getattr(line, '_8').split(',')]
        item = {
            "name": line.name,
            "followres": followres
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

        for empresa in list_empresas[72:]:
            print(empresa)
            if empresa['name'] not in company_igonre:
                input_search = page.locator('//html/body/div[1]/div/div[1]/div/div[2]/div/div/div/input')
                await input_search.press('Backspace')
                await input_search.type(empresa['name'])
                await page.locator('css=.d-print-none.btn.border-0.p-0[role="button"]').click()
                time.sleep(1)

                await page.locator("css=strong.o_kanban_record_title > span", has_text=empresa['name']).click()
                time.sleep(1)


                await page.locator("//html/body/div[1]/div/div/div[2]/div/div[2]/div/div[1]/div/div/div/button").click()
                await page.locator("//html/body/div[2]/div[2]/div[1]/div/a").click()
                time.sleep(1)

                input_followres = page.locator("//html/body/div[2]/div[2]/div[1]/div/div/div/div/main/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div/input")

                for followres in empresa['followres']:
                    await input_followres.type(followres)
                    await input_followres.press('Enter')
                    time.sleep(1)

                await page.locator("//html/body/div[2]/div[2]/div[1]/div/div/div/div/footer/footer/button[1]").click()
                
                time.sleep(1)
                await page.locator("//html/body/div[1]/div/div/div[1]/div/div[1]/div[2]/ol/li[3]/a").click()
                

                time.sleep(2)
        await browser.close()
    



async def init():
    data = await read_xlsx("IMPORTAÇÃO - ODOO - TESTE JHOW (20).xlsx")
    list_companys = await config_list_empresas(data)
    pprint.pprint(list_companys[87:][0])
    # await register_odoo(list_companys)

asyncio.run(init())
