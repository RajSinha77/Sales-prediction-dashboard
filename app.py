import dash  
from dash import dcc, html, Input, Output, State  
import pandas as pd  
import base64  
import io  
from dash.exceptions import PreventUpdate  
from dash import dash_table  
import numpy as np  
  
# Initialize the Dash app  
app = dash.Dash(__name__)  
server = app.server  # Add this line to deploy on Render or similar services  
  
# Siemens colors  
SIEMENS_BLUE = '#009999'  
BG_COLOR = '#F2F2F2'  
  
app.layout = html.Div(style={'backgroundColor': BG_COLOR, 'padding': '20px'}, children=[  
    html.H1('SieSales Opportunity Success Predictor',   
            style={'textAlign': 'center', 'color': SIEMENS_BLUE, 'fontFamily':'Arial'}),  
  
    html.Div([  
        dcc.Upload(  
            id='upload-data',  
            children=html.Div(['Drag and Drop or ', html.A('Select Excel File')]),  
            style={  
                'width': '100%', 'height': '60px', 'lineHeight': '60px',  
                'borderWidth': '2px', 'borderStyle': 'dashed',  
                'borderRadius': '5px', 'textAlign': 'center',  
                'margin': '10px', 'color': SIEMENS_BLUE, 'fontFamily':'Arial'  
            },  
            multiple=False  
        ),  
        html.Div(id='output-data-upload'),  
  
        html.Div(id='loading-output', style={'marginTop':'20px'}),  
        html.Div(id='download-link', style={'marginTop':'20px'})  
    ])  
])  
  
# Dummy ML Prediction function (replace this with your model)  
def run_dummy_ml(df):  
    df['Predicted Outcome'] = np.random.choice(['Win', 'Loss'], size=len(df))  
    return df  
  
@app.callback(  
    Output('output-data-upload', 'children'),  
    Output('loading-output', 'children'),  
    Output('download-link', 'children'),  
    Input('upload-data', 'contents'),  
    State('upload-data', 'filename')  
)  
def update_output(content, filename):  
    if content is None:  
        raise PreventUpdate  
      
    content_type, content_string = content.split(',')  
    decoded = base64.b64decode(content_string)  
    try:  
        if 'xls' in filename:  
            df = pd.read_excel(io.BytesIO(decoded))  
        else:  
            return html.Div(['Only Excel files are accepted.']), None, None  
    except Exception as e:  
        return html.Div([f'There was an error processing this file: {e}']), None, None  
      
    loading_indicator = dcc.Loading(  
        id="loading-1",  
        type="circle",  
        children=html.Div('ML model is running...', style={'color': SIEMENS_BLUE, 'fontSize': '20px'})  
    )  
  
    # Run Dummy ML prediction  
    df_result = run_dummy_ml(df)  
  
    # Create downloadable Excel  
    output = io.BytesIO()  
    df_result.to_excel(output, index=False)  
    output.seek(0)  
    excel_data = base64.b64encode(output.read()).decode()  
  
    download_link = html.A(  
        'Click here to download predicted Excel file',  
        href='data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + excel_data,  
        download='sie_sales_predictions.xlsx',  
        style={  
            'display':'block', 'padding':'10px', 'marginTop':'20px',  
            'backgroundColor': SIEMENS_BLUE, 'color':'white', 'width':'50%',  
            'textAlign':'center','borderRadius':'5px', 'textDecoration':'none',  
            'fontFamily':'Arial'  
        }  
    )  
  
    # Result preview table  
    table_preview = dash.dash_table.DataTable(  
        data=df_result.head(10).to_dict('records'),  
        columns=[{"name": i, "id": i} for i in df_result.columns],  
        style_header={'backgroundColor': SIEMENS_BLUE, 'color':'white'},  
        style_cell={'textAlign':'center', 'fontFamily':'Arial'},  
        page_size=10,  
    )  
  
    return table_preview, loading_indicator, download_link  
  
if __name__ == '__main__':  
    app.run_server(debug=True)  
