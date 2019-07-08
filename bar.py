import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from mpl_toolkits.mplot3d import axes3d
from matplotlib.font_manager import FontProperties
from openpyxl import load_workbook
from pandas import DataFrame


def path_for_wb_to_df(filename, p):
    ws = load_workbook(filename=filename, read_only=True).worksheets[p]
    return DataFrame(ws.values)


def my_plot(df, r_start, r_stop, c_start, c_stop):
    c_start = c_start - 1
    fp = FontProperties(fname=r'C:\WINDOWS\Fonts\YuGothic.ttf', size=30)
    dz = []
    for i in range(c_stop):
        dz.append(df.loc[r_start:r_stop, c_start + i].values)
    print(dz)
    z_pos = df.loc[r_start:r_stop, c_start].values * 0
    x_labels = pd.Index(['-20', '-50', '-80'],
                        dtype='object', name='x')
    y_labels = pd.Index(['60', '80', '100'],
                        dtype='object', name='y')
    x = np.arange(x_labels.shape[0])
    y = np.arange(y_labels.shape[0])
    x_M, y_M = np.meshgrid(x, y, copy=False)
    plt.style.use('seaborn-white')
    ax = plt.figure(figsize=(15, 15)).add_subplot(111, projection='3d')
    # Making the intervals in the axes match with their respective entries
    ax.w_xaxis.set_ticks(x + 0.5 / 2.)
    ax.w_yaxis.set_ticks(y + 0.5 / 2.)
    # Renaming the ticks as they were before
    ax.w_xaxis.set_ticklabels(x_labels)
    ax.w_yaxis.set_ticklabels(y_labels)
    plt.tick_params(labelsize=35)
    # Labeling the 3 dimensions
    ax.set_xlabel('沈み込み量', fontproperties=fp, labelpad=40)
    ax.set_ylabel('荷重', fontproperties=fp, labelpad=40)
    ax.set_zlabel('該当数量', fontproperties=fp, labelpad=40)
    colors = plt.get_cmap("tab10")
    rect = []
    for i in range(c_stop):
        ax.bar3d(x_M.ravel(), y_M.ravel(), z_pos, dx=0.3, dy=0.3, dz=dz[i], color=colors(i))
        z_pos += dz[i]
        rect.append(plt.Rectangle((0, 0), 1, 1, fc=colors(i)))
    ax.legend(rect, df.loc[r_start - 1, c_start:c_start + c_stop - 1], prop=fp,
              loc='upper left')


def main():
    df = path_for_wb_to_df('GP1A173LCS54特作結果.xlsx', 1)
    my_plot(df, 6, 14, 11, 3)
    path = './python/'
    plt.savefig(path + "wbqc.png", format='png', dpi=300)
    my_plot(df, 6, 14, 17, 4)
    plt.savefig(path + "ele.png", format='png', dpi=300)
    my_plot(df, 6, 14, 27, 9)
    plt.savefig(path + "app.png", format='png', dpi=300)
    df = path_for_wb_to_df('GP1A173LCS54特作結果.xlsx', 0)
    my_plot(df, 4, 12, 11, 2)
    plt.savefig(path + "xray.png", format='png', dpi=300)
    my_plot(df, 4, 12, 14, 2)
    plt.savefig(path + "curve.png", format='png', dpi=300)
    my_plot(df, 4, 12, 17, 2)
    plt.savefig(path + "re_app.png", format='png', dpi=300)


main()
