import dash
import dash_core_components as dcc
import dash_html_components as html
from xlrd import open_workbook
from scp import SCPClient


 
import paramiko
import sys

external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']

app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

#all imports are here

#app = dash.Dash(name)

server = app.server

nbytes = 4096
hostname = 'RSCTBDEV1.fyre.ibm.com'
port = 22
username = 'root'
password = 'Mstbsep20!8'
#remotefile="/tmp/MonitoringDailyReport/OutputAppend"+time.strftime('%Y%m%d')+".xls"
#file="OutputAppend"+time.strftime('%Y%m%d')+".xls"


ssh = paramiko.SSHClient()
ssh.load_system_host_keys()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
ssh.connect(hostname, username="root", password="Mstbsep20!8")
print("Connected")
scp = SCPClient(ssh.get_transport())
scp.get('/tmp/MonitoringDailyReport/OutputAppend20190213.xls')
scp.close()

wb = open_workbook('OutputAppend20190213.xls')
createOrder=[]
ScheduledOrder=[]
ReleaseOrder=[]
ShippedOrder=[]
CanceledOrder=[]
ReturnCreated=[]

for sheet in wb.sheets():
    number_of_rows = sheet.nrows
    number_of_columns = sheet.ncols
    time=[]

    items = []

    rows = []
    for col in range(1, number_of_columns):
        values = []
        for row in range(0, number_of_rows):
            value = (sheet.cell(row, col).value)
            try:
                if(row==0):
                    time.append(value)
                if(row==1):
                    createOrder.append(value);
                if(row==2):
                    ScheduledOrder.append(value);
                if(row==3):
                    ReleaseOrder.append(value)
                if (row == 4):
                    ShippedOrder.append(value)
                if(row== 5):
                    CanceledOrder.append(value)
                if(row ==6 ):
                    ReturnCreated.append(value)

               # switch_func(row,value)
               # value = str(int(value))
                #print value
            except ValueError:
                pass
            finally:
                values.append(value)

#print 'Done'







app.layout = html.Div(children=[
    html.H1(children='MS Automation Dash'),
html.Div(id='live-update-text'),
        #dcc.Graph(id='live-update-graph'),
    dcc.Graph(
        id='example-graph',
        figure={
            'data': [
                {'x': time, 'y': createOrder, 'type': 'line', 'name': 'CreateOrder', 'id': 'live-update-graph'},
                {'x': time, 'y': ReleaseOrder, 'type': 'line', 'name': 'ReleaseOrder'},
                {'x': time, 'y': CanceledOrder, 'type': 'line', 'name': 'CanceledOrder'},
                {'x': time, 'y': ReturnCreated, 'type': 'line', 'name': 'ReturnCreated'},

            ],
            'layout': {
                'title': 'MS Automation DashBoard'
            }
        }
    )
])

if __name__ == '__main__':
  app.run_server(debug=True)

