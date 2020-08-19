from jinja2 import Environment, FileSystemLoader
from outlook.email import novo_email
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
import log
import traceback


def folder_check(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)


def figdatab64(plt):
    from io import BytesIO
    figfile = BytesIO()
    plt.savefig(figfile, format='png')
    figfile.seek(0)  # rewind to beginning of file
    import base64
    return base64.b64encode(figfile.getvalue()).decode('utf8')


def img_outlook_b64(path_file, x=155, y=800, dpi=80):
    import matplotlib.image as mpimg
    margin = 0
    xpixels, ypixels = x, y
    figsize = (1 + margin) * ypixels / dpi, (1 + margin) * xpixels / dpi
    fig = plt.figure(figsize=figsize, dpi=dpi)
    ax = fig.add_axes([margin, margin, 1 - 2 * margin, 1 - 2 * margin])
    ax.set_axis_off()
    plt.imshow(mpimg.imread(path_file))
    return figdatab64(plt)


def report_test(env,path_img: str,send_flag=False):
    
    plt.figure(figsize=(1, 1), dpi=96)
    fig, ax = plt.subplots(figsize=(8,3), subplot_kw=dict(aspect="equal"))
    
    #ax.set_position([0.1, 0.05, 0.5, 1])
    x = np.arange(0,4*np.pi,0.2)
    y = 2*np.sin(x)
    ax.plot(x,y)
    res = figdatab64(plt)
    plt.close()

    d = {'col1': [0,1,2], 'col2':[3,4,5]}
    df1 = pd.DataFrame(data=d)
    df2 = pd.DataFrame(data=d)

    template = env.get_template('outlook_test.html')
    header_img = img_outlook_b64(os.path.join(path_img, 'header.png'),y=801, x=59, dpi=96)

    html = template.render(texto1='Lorem ipsum',
        df1=df1,
        df2=df2,
        chart_fig = res,
        texto2='Lorem ipsum',
        texto3='Lorem ipsum',
        header_img=header_img)

    to_list = 'tst@test'
    assunto = 'Lorem ipsum'

    novo_email(html=html, assunto=assunto, to_list=to_list,send=send_flag)
    plt.close('all')


def informes(informe: str, ueps = None, offset_mes: int = 0, send_flag=False):
    logger = log.setup_custom_logger('outlook')
    try:
        # Templates
        file_path = os.path.normpath(os.path.join(os.path.dirname(__file__), 'templates'))
        path_img = file_path.replace('\\', '/')
        dir_jinja = file_path.replace('\\', '//') + '//'
        env = Environment(loader=FileSystemLoader(dir_jinja))

        if informe == 'test':
            report_test(env=env, path_img=path_img, send_flag=send_flag)
    except Exception as e:
        logger.error(traceback.format_exc())
        print(e)


if __name__ == "__main__":
    informes('test')
