def reporter(excelfile):
    __author__ = 'Glezma'
    import webbrowser
    import xlrd
    from flask import Markup

    book = xlrd.open_workbook(excelfile)
    grantitulo = book.sheet_by_index(0).cell(1,2).value
    memotext = book.sheet_by_index(0).cell(2,2).value
    pdftext = book.sheet_by_index(0).cell(3,2).value
    
    table_list =[]
    plot_list =[]
    title_list = []
    description_list = []
 # range(1,book.nsheets)
    for indexsheet in range(1,book.nsheets):
        summary_table, plot_url, titulo ,descripcion = plot_and_tables(excelfile,indexsheet)
        table_list.append(Markup(summary_table))
        plot_list.append(Markup(plot_url))
        title_list.append(Markup(titulo))
        description_list.append(Markup(descripcion))
    # import pdb as pdb
    # pdb.set_trace()
    memotext = Markup(memotext)
    pdftext = Markup(pdftext)
    grantitulo = Markup(grantitulo)
    nlen = len(title_list)
    return(grantitulo, memotext, pdftext, 
            table_list, plot_list, title_list, description_list,nlen)


  # titulo summary_table plot_url descripcion
def plot_and_tables(excelfile,indexsheet):
    import xlrd
    import pandas as pd
    import plotly.plotly as py
    import plotly.graph_objs as pl
    xls = pd.ExcelFile(excelfile)
    book = xlrd.open_workbook(excelfile)
    this_sheet = book.sheet_by_index(indexsheet)
    nrows = int(this_sheet.cell(3, 2).value)
    ncols = int(this_sheet.cell(4, 2).value)
    indexopt = int(this_sheet.cell(7, 0).value)
    incluir_total_footer = int(this_sheet.cell(3, 9).value)
    type_graph_ind = list()
    trace_colors = list()
    for ii in range(1, ncols+1):
        type_graph_ind.append(int(this_sheet.cell(7, ii).value))
        trace_colors.append(this_sheet.cell(8, ii).value)
    titulo = this_sheet.cell(2, 2).value
    descripcion = this_sheet.cell(5,1).value
    typeofgraph = int(this_sheet.cell(6, 2).value)
    if this_sheet.cell(3, 6).value =="a":
        maxvalue ="a"
        minvalue ="a"
    else:
        maxvalue = int(this_sheet.cell(3, 6).value)
        minvalue = int(this_sheet.cell(4, 6).value)
    if indexopt==0:
        data = xls.parse(indexsheet, skiprows=9,parse_cols=ncols, na_values=['NA'])
    else:
        # index_col=["none"]
        data = xls.parse(indexsheet, skiprows=9, parse_cols=ncols, na_values=['NA'])
    print(nrows)
    df = data[:nrows]
    print(df.head())

    print('Llamando bibliotecas')
    #Plot with plotly
    py.sign_in("glezma", "0q6w6pozu7")
    print('Hecho!!')
    print('Procesando graficos')
    listdata = list()
    for count in range(1,ncols+1):
        print(type_graph_ind[count-1])
        print(type_graph_ind)
        if type_graph_ind[count-1]==0:
            df.head()
            plot = pl.Scatter(x=df['Fecha'], y=df.ix[:,count], mode='lines+markers', marker=pl.Marker(size=8), name=df.columns[int(count)])
            listdata.append(plot)
            print(count)
            layout = pl.Layout()
        elif type_graph_ind[count-1]==1:
            df.head()
            print(df['Fecha'])
            if df.ix[:,count].iloc[-1]>=0:
                plot = pl.Bar(x=df['Fecha'], y=df.ix[:,count], name=df.columns[int(count)],yaxis='y1', marker=pl.Marker(
        color=trace_colors[int(count-1)] )  )
            else:
                plot = pl.Bar(x=df['Fecha'], y=df.ix[:,count], name=df.columns[int(count)],yaxis='y2',marker=pl.Marker(
        color=trace_colors[int(count-1)] ))
            listdata.append(plot)
            print(count)
        else:
            plot = pl.Scatter(x=df['Fecha'], y=df.ix[:,count], mode='lines+markers', marker=pl.Marker(size=8, color='rgba(0, 0, 0, 0.95)'), name=df.columns[int(count)],yaxis='y2')
            listdata.append(plot)
            print(count)
            layout = pl.Layout()
        if minvalue!="a":
            layout = pl.Layout(barmode='stack',bargap=0.6,yaxis=pl.YAxis(title='yaxis title',range=[minvalue, maxvalue]),
                               yaxis2=pl.YAxis(title='yaxis title',side='right',overlaying='y',
                                               tickfont=pl.Font(color='rgb(1, 1, 1)'),range=[minvalue, maxvalue]))
        else:
            layout = pl.Layout()
            #layout = pl.Layout(barmode='stack',yaxis=pl.YAxis(title='yaxis title'),yaxis2=pl.YAxis(title='yaxis title',side='right',overlaying='y'))
    print('Hecho!!')
    pdata = pl.Data(listdata)
    fig = pl.Figure(data=pdata, layout=layout)
    print('Intentando conexion remota...')
    plot_url = py.plot(fig, filename='Repjs_'+ str(indexsheet), auto_open=False)
    plot_url = plot_url + '.embed'
    # plot_url = 'https://plot.ly/~glezma/271.embed'

    df1=df.set_index('Fecha').T
    summary_table = df1 .to_html()    .replace('<table border="1" class="dataframe">', '<table class="display", align = "center", style="width:100%;">')  # use bootstrap styling
    summary_table = summary_table    .replace('<tr style="text-align: right;">', '<tr>')  # use bootstrap styling
    if incluir_total_footer != 0:
        lastindex = df1.index.values[-1]
        toreplace = '''<tr>\n      <th>''' + lastindex
        toplace = '<tfoot>\n <tr>\n      <th>' + lastindex
        summary_table = summary_table    .replace(toreplace, toplace)  # use bootstrap styling
        toreplace ='</tbody>'
        toplace ='</tfoot></tbody>'
        summary_table = summary_table    .replace(toreplace, toplace)
    return (summary_table, plot_url, titulo ,descripcion)
  

