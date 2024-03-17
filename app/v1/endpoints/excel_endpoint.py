from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from openpyxl import load_workbook
from jinja2 import Template
import pandas as pd
import os
import uuid
import shutil

router = APIRouter()
EXCEL_WORKSPACE = os.path.abspath("app\\v1\\workspace\\excel\\")

class ExcelFilePath(BaseModel):
    excel_file_1_path: str
    excel_file_1_sheet_number: int = 1
    excel_file_2_path: str
    excel_file_2_sheet_number: int = 1

@router.post("/generate_url_for_excel_doc")
async def generate_url(file_paths: ExcelFilePath):
    document_1 = file_paths.excel_file_1_path
    document_1_sheet_number = file_paths.excel_file_1_sheet_number
    document_2 = file_paths.excel_file_2_path
    document_2_sheet_number = file_paths.excel_file_2_sheet_number
    
    """validate the document paths"""
    if not os.path.isfile(document_1):
        raise HTTPException(status_code=404, detail=f"{document_1} not found.")
    if not os.path.isfile(document_2):
        raise HTTPException(status_code=404, detail=f"{document_2} not found.")
    
    """validate the document format. XLSX is the valid format"""
    if not validate_xlsx_format(document_1, document_2):
        raise HTTPException(status_code=400, detail="Invalid document format")
    
    """validate the sheet number"""
    if not validate_excel_sheet_number(document_1, document_1_sheet_number, document_2, document_2_sheet_number):
        raise HTTPException(status_code=400, detail="Invalid docment sheet number")
    
    """check if the given document is not blank"""
    if is_excel_sheet_blank(document_1, document_1_sheet_number, document_2, document_2_sheet_number):
        raise HTTPException(status_code=400, detail="Document is blank")

    """Create workspace for the session request"""
    get_session_id = create_workspace_for_session()
    EXCEL_WORKSPACE_SESSION = os.path.join(EXCEL_WORKSPACE, get_session_id)

    """Copy the documents to session workspace"""
    if not copy_document_to_session_workspace(document_1, document_2, EXCEL_WORKSPACE_SESSION):
        raise HTTPException(status_code=500, detail="Error while copying the docuemnts to session workspace")
    
    """read the excel documents using Pandas"""
    new_doc1_path = os.path.join(EXCEL_WORKSPACE_SESSION, os.path.basename(document_1))
    new_doc2_path = os.path.join(EXCEL_WORKSPACE_SESSION, os.path.basename(document_2))

    df1 = pd.read_excel(new_doc1_path, sheet_name=document_1_sheet_number -1)
    df2 = pd.read_excel(new_doc2_path, sheet_name=document_2_sheet_number -1)

    """ensure both the dataframes have the same columns"""
    columns = list(set(df1.columns) | set(df2.columns))
    df1 = df1.reindex(columns=columns)
    df2 = df2.reindex(columns=columns)

    """find the length of each dataframe"""
    len_df1 = len(df1)
    len_df2 = len(df2)

    """check which DataFrame has more rows """
    if len_df1 > len_df2:
        """append missing rows in df2 with None values"""
        missing_rows = len_df1 - len_df2
        df2 = pd.concat([df2, pd.DataFrame([[None]*len(columns)]*missing_rows, columns=columns)], ignore_index=True)
    elif len_df2 > len_df1:
        """Append missing rows in df1 with None values"""
        missing_rows = len_df2 - len_df1
        df1 = pd.concat([df1, pd.DataFrame([[None]*len(columns)]*missing_rows, columns=columns)], ignore_index=True)
    
    """heightlight the rows"""
    diff_df = df1.compare(df2)
    highlighted_rows = diff_df.index.tolist()

    """get the indices of all rows that are different or present only in one dataframe"""
    highlighted_rows = list(set(diff_df.index).union(set(df1.index).symmetric_difference(set(df2.index))))

    """convert dataframes to dictionary"""
    table1 = df1.to_dict(orient='records')
    table2 = df2.to_dict(orient='records')

    """update table1"""
    new_table_1 = [convert_keys_to_strings(item) for item in table1]
    new_table_1 = update_data_with_default(new_table_1, "N/A")

    """update table2"""
    new_table_2 = [convert_keys_to_strings(item) for item in table2]
    new_table_2 = update_data_with_default(new_table_2, "N/A")

    """write comparision result"""
    generate_html_file(EXCEL_WORKSPACE_SESSION, "Contentverse Excel Document Comparision", document_1, 
                       document_1_sheet_number, new_table_1,
                        document_2, document_2_sheet_number, new_table_2, 
                       highlighted_rows)

    return {"message": "File Found"}

"""func: convert integer keys to string"""
def convert_keys_to_strings(dictionary):
    updated_dict = {}
    for key, value in dictionary.items():
        if isinstance(key, int):
            updated_dict["ID"] = value
        else:
            updated_dict[key] = value
    return updated_dict

"""func: update data with default value"""
def update_data_with_default(data_list, default_value):
    updated_data = []
    for item in data_list:
        updated_item = {key: value if value is not None else default_value for key, value in item.items()}
        updated_data.append(updated_item)
    return updated_data

"""func: validate the xlsx format document"""
def validate_xlsx_format(file1_path: str, file2_path: str) -> bool:
    """check if the document extentions endswith .xlsx"""
    if not file1_path.endswith(".xlsx") or not file2_path.endswith(".xlsx"):
        return False
    """check if the document can be opened"""
    try:
        load_workbook(file1_path)
        load_workbook(file2_path)
        return True
    except Exception as error:
        return False

"""func: validate the excel document sheet number"""
def validate_excel_sheet_number(file1_path: str, file1_sheet_number: int, file2_path: str, file2_sheet_number: int) -> bool:
    """load the workbook"""
    workbook_1 = load_workbook(file1_path)
    workbook_2 = load_workbook(file2_path)

    """check the sheet number is valid"""
    if file1_sheet_number < 1 or file1_sheet_number > len(workbook_1.sheetnames):
        return False
    if file2_sheet_number < 1 or file2_sheet_number > len(workbook_2.sheetnames):
        return False
    return True

"""func: excel document is blank"""
def is_excel_sheet_blank(file1_path: str, file1_sheet_number: int, file2_path: str, file2_sheet_number: int) -> bool:
    """load the workbook"""
    workbook_1 = load_workbook(file1_path)
    workbook_2 = load_workbook(file2_path)

    """get the index by index"""
    doc_1_sheet = workbook_1.worksheets[file1_sheet_number - 1]
    doc_2_sheet = workbook_2.worksheets[file2_sheet_number - 1]

    """doc1 sheet"""
    for row in doc_1_sheet.iter_rows():
        for cell in row:
            if cell.value:
                return False
    """doc2 sheet"""
    for row in doc_2_sheet.iter_rows():
        for cell in row:
            if cell.value:
                return False
    return True

"""func: workspace for new session request"""
def create_workspace_for_session():
    session_id = str(uuid.uuid4())
    session_folder = os.path.join(EXCEL_WORKSPACE, session_id)
    os.makedirs(session_folder)
    return session_id


"""func: copy documents to session workspace"""
def copy_document_to_session_workspace(file1_path: str, file2_path, session_workspace_path: str):
    try:
        shutil.copy(file1_path, session_workspace_path)
        shutil.copy(file2_path, session_workspace_path)
        return True
    except Exception as error:
        return False


from jinja2 import Template

def generate_html_file(session_path, title, file1, file1_sheet_number,
                        data1, file2, file2_sheet_number,
                          data2, highlighted_rows):
    # Jinja2 template string
    template_str = '''<!DOCTYPE html>
<html>
<head>
    <title>{{ title }}</title>
    <!-- Bootstrap CSS -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous">
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.7.1/jquery.min.js"></script>
    <style>
        .table-responsive {
            overflow-x: unset;
        }
        .table-responsive table {
            width: 100%;
        }
        .table-container {
            display: flex;
        }
        .table-container .table-responsive {
            flex: 0 0 auto;
            margin-right: 10px;
        }
        .square-badge {
            border-radius: 0;
        }
        .divScrollDiv {
            display: inline-block;
            width: 100%;
            border: 1px solid black;
            height: 94vh;
            overflow: scroll;
        }
        .tableNoScroll {
            overflow: hidden;
        }
    </style>
    <script>
        $(document).ready(function () {
            var target_sec = $("#divFixed");
            $("#divLista").scroll(function () {
                target_sec.prop("scrollTop", this.scrollTop)
                    .prop("scrollLeft", this.scrollLeft);
            });
            var target_first = $("#divLista");
            $("#divFixed").scroll(function () {
                target_first.prop("scrollTop", this.scrollTop)
                    .prop("scrollLeft", this.scrollLeft);
            });
        });
    </script>
</head>
<body>
    <div class="col-lg mx-auto p-1 py-md-1">
        <header class="d-flex align-items-center pb-1">
            <a href="#" class="d-flex align-items-center text-dark text-decoration-none">
                <img src="/static/images/logo.jpeg" width="32" height="32" class="p-1">
                <span class="fs-6">Document Comparison and Analysis (Demo)</span>
            </a>
        </header>
        <div class="container-fluid">
            <div class="row">
                <div class="col divScrollDiv border" id="divFixed">
                    <div class="table-container">
                        <!-- First Table -->
                        <div class="table table-responsive mt-2">
                            <span class="badge bg-primary square-badge">Excel Document Path</span><span class="badge bg-success square-badge">{{ file1 }}</span><br>
                            <span class="badge bg-primary square-badge">Excel Sheet Number</span><span class="badge bg-success square-badge">{{ file1_sheet_number }}</span>
                            <table class="table-sm table-bordered mt-2">
                                <thead class="table-dark">
                                    <tr>
                                        {% for i in data1[0].keys() %}
                                            <th>{{ i }}</th>
                                        {% endfor %}
                                    </tr>
                                </thead>
                                <tbody>
                                    {% for row in data1 %}
                                        <tr {% if loop.index0 in highlighted_rows %}class="table-warning"{% endif %}>
                                            {% for value in row.values() %}
                                                <td>{{ value }}</td>
                                            {% endfor %}
                                        </tr>
                                    {% endfor %}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
                <div class="col divScrollDiv border" id="divLista">
                    <div class="table-container">
                        <div class="table table-responsive mt-2">
                            <!-- Second Table -->
                            <span class="badge bg-primary square-badge">Excel Document Path</span><span class="badge bg-success square-badge">{{ file2 }}</span><br>
                            <span class="badge bg-primary square-badge">Excel Sheet Number</span><span class="badge bg-success square-badge">{{ file2_sheet_number }}</span>
                            <table class="table-sm table-bordered mt-2">
                                <thead class="table-dark">
                                    <tr>
                                        {% for i in data2[0].keys() %}
                                            <th>{{ i }}</th>
                                        {% endfor %}
                                    </tr>
                                </thead>
                                <tbody>
                                    {% for row in data2 %}
                                        <tr {% if loop.index0 in highlighted_rows %}class="table-warning"{% endif %}>
                                            {% for value in row.values() %}
                                                <td>{{ value }}</td>
                                            {% endfor %}
                                        </tr>
                                    {% endfor %}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <!-- Bootstrap JS (optional) -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/js/bootstrap.bundle.min.js" integrity="sha384-MrcW6ZMFYlzcLA8Nl+NtUVF0sA7MsXsP1UyJoMp4YLEuNSfAP+JcXn/tWtIaxVXM" crossorigin="anonymous"></script>
</body>
</html>'''

    # Create a Jinja2 template
    template = Template(template_str)

    # Render the template with the provided data
    rendered_html = template.render(
        title=title,
        file1=file1,
        file1_sheet_number=file1_sheet_number,
        data1=data1,
        file2=file2,
        file2_sheet_number=file2_sheet_number,
        data2=data2,
        highlighted_rows=highlighted_rows
    )

    # Write the rendered HTML to file
    with open(f"{session_path}/comparision_result.html", "w") as file:
        file.write(rendered_html)
