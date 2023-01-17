#%% Importar librerías
import xlwings as xw
import numpy as np
import pandas as pd
from Code.Funciones import Rs
from Code.Funciones import Bo
from Code.Funciones import Bg

#%% Etiquetas
INPUTS = "Input"
RESULTS = "Results"
Rs_ = "Rs"
Rsb_ = "Rsb"
Bo_ = "Bo"
Pb_ = "Pb"
Uo_ = "Uo"
API = "°API"
G_GAS = "ϒgas"
G_OIL = "ϒoil"

#%% Main
def main():
    wb = xw.Book.caller()
    sheet_input = wb.sheets[INPUTS]
    sheet_results = wb.sheets[RESULTS]
#Datos
    v_P = sheet_input["C7:C19"].value
    num_datos = len(v_P)
    v_Np = sheet_input["D7:D17"].value
    v_Wp = sheet_input["E7:E17"].value
    Pb_ = sheet_input["E25"].value
    API = sheet_input["E26"].value
    T = sheet_input["E27"].value
    G_GAS = sheet_input["E28"].value
    G_OIL = sheet_input["E29"].value
    Z = sheet_input["E30"].value
    Swi = sheet_input["E31"].value
    Bw = sheet_input["E32"].value
    Cw = sheet_input["E33"].value
    Cf = sheet_input["E34"].value
    corr_Rs = sheet_input["I5"].value
    corr_Bo = sheet_input["J5"].value
    df_data = sheet_input["C6"].options(pd.DataFrame, expand = "table", index = False).value
    df_results = sheet_input["I6"].options(pd.DataFrame, expand="table", index=False).value

#Solubidad del gas (Rs)

    Rs_Corr = []
    v_Psat = []
    v_API = []
    v_Temp = []
    v_G_gas = []
    v_G_oil = []
    for i in range(num_datos):
        Rs_Corr.append(corr_Rs)
        v_Psat.append(Pb_)
        v_API.append(API)
        v_Temp.append(T)
        v_G_gas.append(G_GAS)
        v_G_oil.append(G_OIL)


    Parametros_Rs = pd.DataFrame(
        {"Correlacion": Rs_Corr, "Presion": v_P, "Presion_Burb": v_Psat, "API": v_API,
         "Temperatura": v_Temp, "G_gas": v_G_gas, "G_oil": v_G_oil})

    v_Rs = []
    for i in range(num_datos):
        Rs_resul = Rs(*(
            Parametros_Rs.iloc[i, 0], Parametros_Rs.iloc[i, 1], Parametros_Rs.iloc[i, 2], Parametros_Rs.iloc[i, 3],
            Parametros_Rs.iloc[i, 4],
            Parametros_Rs.iloc[i, 5], Parametros_Rs.iloc[i, 6]))
        v_Rs.append(Rs_resul)
    #sheet_results["B5"].options(pd.DataFrame).value = Parametros_Rs
    sheet_input["I7"].options(transpose = True).value = v_Rs

#Factor volumétrico de formación del crudo (Bo)

    Bo_Corr = []
    for i in range(num_datos):
        Bo_Corr.append(corr_Bo)

    Params_Bo = pd.DataFrame(
        {"Correlacion": Bo_Corr, "Presion": v_P, "Presion_burb": v_Psat, "Rs": v_Rs,
         "Rsb": v_Rs, "G_gas": v_G_gas, "G_oil": v_G_oil, "Temperatura": v_Temp,
         "API": v_API})
    v_Bo = []
    for i in range(num_datos):
        Bo_resul = Bo(*(
            Params_Bo.iloc[i, 0], Params_Bo.iloc[i, 1], Params_Bo.iloc[i, 2], Params_Bo.iloc[i, 3],
            Params_Bo.iloc[i, 4],
            Params_Bo.iloc[i, 5], Params_Bo.iloc[i, 6], Params_Bo.iloc[0, 7], Params_Bo.iloc[0, 8]))
        v_Bo.append(Bo_resul)
    sheet_input["J7"].options(transpose=True).value = v_Bo


if __name__ == "__main__":
    xw.Book("Controlador.xlsx").set_mock_caller()
    main()