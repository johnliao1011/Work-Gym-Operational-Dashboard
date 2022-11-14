import dash
import dash_daq as daq
import dash_core_components as dcc
import dash_html_components as html
import dash_bootstrap_components as dbc
from dash.dependencies import Input, Output, State
from datetime import datetime
import pandas as pd
import numpy as np
import datetime
import plotly.express as px
import plotly.graph_objects as go

# Start the app
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.CYBORG])
server = app.server

# function
## performance metric
def performance(data, para, start_date, end_date, location):
    objects = data[para].unique()
    result = {}
    

    for object in objects:
        if isinstance(location, list)==True:
            if para is None:
                df = data[(data["日期"]>= start_date) & (data["日期"]<= end_date) & (data["地點"].isin(location))]
            else:
                df = data[(data["日期"]>= start_date) & (data["日期"]<= end_date) & (data["地點"].isin(location)) & (data[para]== object)]
        else:
            if para is None:
                df = data[(data["日期"]>= start_date) & (data["日期"]<= end_date) & (data["地點"] == location)]
            else:
                df = data[(data["日期"]>= start_date) & (data["日期"]<= end_date) & (data["地點"] == location) & (data[para]== object)]

        classes = df["日期"].count()
        students = df["人數"].sum()
        avg_students = "{:.1f}".format((students/classes) if (pd.isna(float(students/classes))== False) else 0)
        prc_students = "{:.1f}%".format((students/(classes*20)*100) if (pd.isna(float((students/(classes*20)*100)))== False) else 0)
        result[object]=[classes, students, avg_students, prc_students]
    
    return result


# Import data
df = pd.read_excel("統計表.xlsx")
df.insert(2,"Datetime", pd.to_datetime(df["日期"].astype(str)+" "+ df["時間"].astype(str)))


df["時間"] = pd.to_datetime(df['時間'], format='%H:%M:%S').dt.hour
df["人數"] = pd.to_numeric(df["人數"])

df.insert(3,"WEEKDAY", df["日期"].dt.dayofweek)
df.insert(4, "平假日", "平日")
df.loc[df['WEEKDAY'].isin([5, 6]), '平假日'] = "假日"
df.insert(5, "時段", "早上")
df.loc[df['時間'].isin(range(13,19)), '時段'] = "下午"
df.loc[df['時間'].isin(range(19,25)), '時段'] = "晚上"
df.drop("時間", inplace = True, axis = 1)

# item = performance(df, "項目", start_date, end_date, location)
# item = list(item.items())

# Build component parts
## spinner on the card
div_alert = dbc.Spinner(html.Div(id="alert-msg"))

## controls
date_input = dbc.FormGroup(
    [
        dbc.Label("時間區間"),
        dcc.DatePickerRange(
            id='date-pick-range',
            start_date = min(df.日期).to_pydatetime().date(),
            end_date = max(df.日期).to_pydatetime().date(),
            calendar_orientation='horizontal'
        ),
    ]
)

location_input = dbc.FormGroup(
    [
        dbc.Label("營運地點"),
        dcc.Dropdown(id="location", 
                 options=[{"label":location, "value":location} for location in df["地點"].unique()],
                 value = [df["地點"][0]],
                 multi = True),
    ]
)

search_button = dbc.FormGroup(
    [
        dbc.Button("查詢", color="primary", id="search", block = True, size="lg"),
    ]
)

controls = dbc.Form([date_input, location_input, search_button])

# card component
cards = [
    dbc.Card(
        [
            html.H5("課堂數", className="card-text"),
            html.H2(id= "classes", className="card-title"),
            dbc.CardFooter(id= "dow_classes"),
        ],
        body=True,
        color="info",
        inverse=True,
        style={"text-align":"center"},
        outline=True,
    ),
    dbc.Card(
        [
            html.H5("學生數", className="card-text"),
            html.H2(id= "students", className="card-title"),
            dbc.CardFooter(id= "dow_students"),
        ],
        body=True,
        color="success",
        inverse=True,
        style={"text-align":"center"},
        outline=True,
    ),
    dbc.Card(
        [
            html.H5("平均每堂人數", className="card-text"),
            html.H2(id= "avg_students", className="card-title"),
            dbc.CardFooter(id= "dow_avg_students"),
        ],
        body=True,
        color="primary",
        inverse=True,
        style={"text-align":"center"},
        outline=True,
    ),
    dbc.Card(
        [
            html.H5("滿堂率", className="card-text"),
            html.H2(id= "prc_students", className="card-title"),
            dbc.CardFooter(id= "dow_prc_students"),
        ],
        body=True,
        color="danger",
        inverse=True,
        style={"text-align":"center"},
        outline=True,
    ),
]

gauges = [
    daq.Gauge(
        id='普拉',
        showCurrentValue=True,
        units= "人/堂",
        value= 5,
        label= {
            "label":'普拉',
            "style":{
                "color":"#FFFFFF",
                'fontSize': 20
            }
        },
        color = "#F87F0C",
        max=20,
        min=0,
        size=200,
    ),
    daq.Gauge(
        id='TRX',
        showCurrentValue=True,
        units="人/堂",
        value= 5,
        label= {
            "label":'TRX',
            "style":{
                "color":"#FFFFFF",
                'fontSize': 20
            }
        },
        color = "#F87F0C",
        max = 20,
        min = 0,
        size=200,
    ),
    daq.Gauge(
        id='瑜珈',
        showCurrentValue=True,
        units="人/堂",
        value= 5,
        label= {
            "label":'瑜珈',
            "style":{
                "color":"#FFFFFF",
                'fontSize': 20
            }
        },
        color = "#F87F0C",
        max=20,
        min=0,
        size=200,
    ),
    daq.Gauge(
        id='壺鈴',
        showCurrentValue=True,
        units="人/堂",
        value=5,
        label= {
            "label":'壺鈴',
            "style":{
                "color":"#FFFFFF",
                'fontSize': 20
            }
        },
        color = "#F87F0C",
        max=20,
        min=0,
        size=200,
    ),
    daq.Gauge(
        id='拳擊',
        showCurrentValue=True,
        units="人/堂",
        value= 5,
        label= {
            "label":'拳擊',
            "style":{
                "color":"#FFFFFF",
                'fontSize': 20
            }
        },
        color = "#F87F0C",
        max=20,
        min=0,
        size=200,
    ),
    daq.Gauge(
        id='舞',
        showCurrentValue=True,
        units="人/堂",
        value=5,
        label= {
            "label":'舞',
            "style":{
                "color":"#FFFFFF",
                'fontSize': 20
            }
        },
        color = "#F87F0C",
        max=20,
        min=0,
        size=200,
    ),
]

# Define Layout
app.layout = dbc.Container(
    fluid = True,
    children = [
        html.H2("健身房營運儀表板"),
        html.Hr(),
        dbc.Row(  
            [
                dbc.Col([dbc.Card(controls, body=True),div_alert], md = 3),
                dbc.Col(dbc.Row([dbc.Col(cards[0]),dbc.Col(cards[1]),dbc.Col(cards[2]),dbc.Col(cards[3])])),
            ],
        ),
        html.Br(),
        dbc.Row(
            [
                dbc.Col(gauges[0], md=2), dbc.Col(gauges[1], md=2), dbc.Col(gauges[2], md=2), dbc.Col(gauges[3], md=2), dbc.Col(gauges[4], md=2), dbc.Col(gauges[5], md=2)
                            
            ], align="center",
        ),
        html.H2("上課趨勢"),
        html.Hr(),
        dbc.Row(
            [
                dbc.Col([html.H5("每日學生上課狀況", style = {"text-align":"center"}), 
                         dcc.Graph(id="all-line-chart", 
                                   className="svg-container",
                                   style={"height": "100vh"},
                                  )
                        ]
                       ),
                dbc.Col([html.H5("每日教練教授人數", style = {"text-align":"center"}),
                         dcc.Graph(id="coach-line-chart", 
                                        className="svg-container", 
                                        style={"height": "100vh"},
                                  )
                        ]
                       ),
            ],   
        ),
        html.Br(),
        html.H2("項目分布"),
        html.Hr(),
        dbc.Row(
            [
                dbc.Col([html.H5("各項目上課人數分布", style = {"text-align":"center"}), dcc.Graph(id="pie-chart", className="svg-container", style={"height": 400})]),
                dbc.Col([html.H5("各時段上課人數分布", style = {"text-align":"center"}), dcc.Graph(id="period-bar", className="svg-container", style={"height": 400})]),
                dbc.Col([html.H5("各教授上課人數分布", style = {"text-align":"center"}), dcc.Graph(id="coach-bar", className="svg-container", style={"height": 400})]),
            ],
        
        ),
    ],
    style = {"margin":"auto", "height": "100vh"},
)

@app.callback(
    [

        Output('classes', "children"),
        Output('students', "children"),
        Output('avg_students', "children"),
        Output('prc_students', "children"),
        Output('dow_classes', "children"),
        Output('dow_students', "children"),
        Output('dow_avg_students', "children"),
        Output('dow_prc_students', "children"),
        Output("普拉", "value"),
        Output("TRX", "value"),
        Output("瑜珈", "value"),
        Output("壺鈴", "value"),
        Output("拳擊", "value"),
        Output("舞", "value"),
        Output("pie-chart", "figure"),
        Output("all-line-chart", "figure"),
        Output("coach-bar", "figure"),
        Output("coach-line-chart", "figure"),
        Output("period-bar", "figure"),
        
    ],

    [Input('search', 'n_clicks')],
    [
        State("date-pick-range", "start_date"),
        State("date-pick-range", "end_date"),
        State("location", "value"),
        
    ],
)

def update_output(n_clicks, start_date, end_date, location):
    # overall statistical values
    classes = df[(df["日期"]>= start_date) & (df["日期"]<= end_date)& (df["地點"].isin(location))]["日期"].count()
    students = df[(df["日期"]>= start_date) & (df["日期"]<= end_date)& (df["地點"].isin(location))]["人數"].sum()
    avg_students = "{:.1f}".format((students/classes) if (pd.isna(float(students/classes))== False) else 0)
    prc_students = "{:.1f}%".format((students/(classes*20)*100) if (pd.isna(float((students/(classes*20)*100)))== False) else 0)
    
    # weekend
    weekend_classes = df[(df["平假日"]=="假日") & (df["日期"]>= start_date) & (df["日期"]<= end_date)& (df["地點"].isin(location))]["日期"].count()
    weekend_students = df[(df['平假日']=="假日") & (df["日期"]>= start_date) & (df["日期"]<= end_date)& (df["地點"].isin(location))]["人數"].sum()
    weekend_avg_students = "{:.1f}".format((weekend_students/weekend_classes) if (pd.isna(float(weekend_students/weekend_classes))== False) else 0)
    weekend_prc_students = "{:.0f}%".format((weekend_students/(weekend_classes*20)*100) if (pd.isna(float((weekend_students/(weekend_classes*20)*100)))== False) else 0)
    
    # weekday
    weekday_classes = df[(df['平假日']=="平日") & (df["日期"]>= start_date) & (df["日期"]<= end_date)& (df["地點"].isin(location))]["日期"].count()
    weekday_students = df[(df['平假日']=="平日") & (df["日期"]>= start_date) & (df["日期"]<= end_date)& (df["地點"].isin(location))]["人數"].sum()
    weekday_avg_students = "{:.1f}".format((weekday_students/weekday_classes) if (pd.isna(float(weekday_students/weekday_classes))== False) else 0)
    weekday_prc_students = "{:.0f}%".format((weekday_students/(weekday_classes*20)*100) if (pd.isna(float((weekday_students/(weekday_classes*20)*100)))== False) else 0)
    
    dow_classes = html.H6("平日:{} | 假日:{}".format(weekday_classes, weekend_classes))
    dow_students = html.H6("平日:{} | 假日:{}".format(weekday_students, weekend_students))
    dow_avg_students = html.H6("平日:{}  |  假日:{}".format(weekday_avg_students, weekend_avg_students))
    dow_prc_students = html.H6("平日:{}  |  假日:{}".format(weekday_prc_students, weekend_prc_students))

    # item
    item = performance(df, "項目", start_date, end_date, location)
    item = list(item.items())

    df2 = df[(df["日期"]>= start_date) & (df["日期"]<= end_date)& (df["地點"].isin(location))]
    
    # pie chart
    df_item = df2.groupby("項目", as_index=False)["人數"].sum()
    pie_fig = go.Figure(go.Pie(labels=df_item["項目"], values=df_item["人數"], hole=.5))
    pie_fig.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(
        color="#FFFFFF"
        )
    )
    
    # all line chart
    all_line_fig = px.bar(df2, x="日期", y="人數", color="項目")
    df3 = df2.groupby("日期", as_index=False)["人數"].sum()
    all_line_fig.add_trace(go.Scatter(x=df3["日期"], y=df3["人數"], name='總人數', text = df3["人數"],textposition='top right'))
    all_line_fig.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(
        color="#FFFFFF"
        )
    )
    all_line_fig.update_xaxes(showline=False)
    all_line_fig.update_yaxes(showline=False)

    
    # coach bar chart 
    name_df = df2.groupby(["教練",'項目'], as_index=False)["人數"].sum().sort_values("人數",ascending = False)
    name_bar = px.bar(name_df, x="人數", y="教練", orientation='h', color = "項目", text = "人數" )
    name_bar.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(
        color="#FFFFFF"
        )
    )
    
    # coach line chart
    coarch_line_df = df2[["日期", "教練","人數"]].groupby(["日期", "教練"], as_index = False).sum()
    coach_fig = px.line(coarch_line_df, x="日期", y="人數", color='教練')
    coach_fig.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(
        color="#FFFFFF"
        )
    )
    coach_fig.update_xaxes(showline=False, showgrid=False, zeroline=True)
    
    # period bar chart
    period_df = df2.groupby(["時段",'項目'], as_index=False)["人數"].sum().sort_values("人數",ascending = False)
    period_bar = px.bar(period_df, x="人數", y="時段", orientation='h', color = "項目", text = "人數")
    period_bar.update_layout(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(
        color="#FFFFFF"
        )
    )
    
    return classes, students, avg_students, prc_students, dow_classes, dow_students, dow_avg_students, dow_prc_students, float(item[0][1][2]), float(item[1][1][2]), float(item[2][1][2]), float(item[3][1][2]), float(item[4][1][2]), float(item[5][1][2]), pie_fig, all_line_fig, name_bar, coach_fig, period_bar



# Run the App
if __name__ == "__main__":
    app.run_server(debug=True, use_reloader=False)