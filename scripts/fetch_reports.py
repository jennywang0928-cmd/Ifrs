# 用瀏覽器從證交所 ESG InfoHub 抓永續報告書申報資料（含聯絡資訊）。用法: python3 fetch_reports.py <輸出.json> <年度,如 2025>
import asyncio, json, os, sys
from playwright.async_api import async_playwright
OUT=sys.argv[1]; YEARS=[int(y) for y in sys.argv[2].split(',')]
JS='''async ([y,m]) => {
 const H={'Accept':'application/json','Content-Type':'application/json'};
 const c=await (await fetch('/api/api/MopsSustainReport/industryAndCompanyCode?marketType='+m,{headers:H})).json();
 const body={year:y,marketType:m,industryNameList:c.data.industryNameList.map(x=>x.industryName),companyCodeList:c.data.companyNameList.map(x=>x.companyCode)};
 const r=await fetch('/api/api/MopsSustainReport/data',{method:'POST',headers:H,body:JSON.stringify(body)});
 return {status:r.status, text:await r.text()};
}'''
async def main():
    async with async_playwright() as p:
        b=await p.chromium.launch(executable_path='/opt/pw-browsers/chromium-1194/chrome-linux/chrome', proxy={'server':os.environ['HTTPS_PROXY']})
        pg=await (await b.new_context(locale='zh-TW')).new_page()
        await pg.goto('https://esggenplus.twse.com.tw/inquiry/report', wait_until='networkidle', timeout=90000)
        res={}
        for y in YEARS:
          for m in (0,1):
            r=await pg.evaluate(JS,[y,m])
            print(y,m,r['status'],r['text'][:300].replace('\n',' '))
            try: res[f'{y}_{m}']=json.loads(r['text'])
            except Exception: pass
        json.dump(res,open(OUT,'w'),ensure_ascii=False)
        await b.close()
asyncio.run(main())
