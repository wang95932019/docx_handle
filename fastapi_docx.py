import uvicorn
from fastapi import FastAPI
from fastapi.responses import JSONResponse
from fastapi.responses import RedirectResponse
from pydantic import BaseModel
from typing import Union
import handle_docx
from pydocx import PyDocX
import config
import handle_excel


class MainItem(BaseModel):
    path: str       # 文件路径
    college: str    # 学校编号。具体对应关系见配置文件，config.py。
    type: str = 'rcpy'       # 文档类型。如：rcpy(人才培养)
    depth: Union[int, None] = 99999


class DocxToHtmlItem(BaseModel):
    path: str


app = FastAPI()


@app.get('/')
def index():
    return RedirectResponse('/docs')


@app.api_route('/handle_docx', methods=['POST'])
async def handle_docx_main(item: MainItem):
    tmp = config.college_relationship[item.college]
    docx = config.college_relationship[item.college][item.type]

    result_code, text_res, res_tables_list, res_images_list, result_key_list, result_table_key_list = handle_docx.main(
        item.path, docx.rules, item.depth, docx.key_rules, docx.new_rules, docx.tables_key_rules)
    content = {
        "msg": "true",
        "code": 200,
        "text": text_res,
        "tables": res_tables_list,
        "images": res_images_list
    }
    if result_key_list:
        content["key_list"] = result_key_list
    if result_table_key_list:
        content["table_key_list"] = result_table_key_list
    if result_code == 404:
        content["msg"] = "false"
        content["code"] = 200
    return JSONResponse(content=content, status_code=200)


@app.api_route('/handle_excel', methods=['POST'])
async def handle_excel_main(item: MainItem):
    course, hours_sum, practice, credit_statistics, commits = handle_excel.main(item.path)
    content = {
        "msg": "true",
        "code": 200,
        "course": course,
        "hours_sum": hours_sum,
        "practice": practice,
        "credit_statistics": credit_statistics,
        "commits": commits
    }

    return JSONResponse(content=content, status_code=200)


# @app.api_route('/handle', methods=['POST'])
# async def handle_docx_main(item: MainItem):

@app.api_route('/docx_to_html', methods=['POST'])
async def docx_to_html(item: DocxToHtmlItem):
    html = PyDocX.to_html(item.path)
    content = {
        "msg": "true",
        "code": 200,
        "html": html
    }
    return JSONResponse(content=content, status_code=200)


if __name__ == '__main__':
    # print(assist_rules[1:2])
    uvicorn.run(app=app, host="127.0.0.1", port=8000)
