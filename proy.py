import math
import pandas as pd
#df = pd.read_excel('PROYE.xlsx')
def lambda_handler (event, context):
    excel_file = 'PROYE.xlsx'
    body = event["body"]

    vm=body["vm"]
    ke=body["ke"]
    h=body["h"]
    vm_um_values = {115: 123, 230:245, 500: 550, 13.8:17.5}
    um = vm_um_values[vm]
    ######################################### calcula el cov y tov
    COV=um/(3**(1/2))
    TOV=ke*COV
    ######################################### hallar el Rn
    R0=COV/0.8
    RE=TOV/1.1
    if (R0>RE or R0==RE ):
        Rn=R0
    if R0<RE :
           Rn=RE
    Rnn=int(Rn/3)*3

    #################################3########## seleccionar el NPR Y NPM del catalogo
    df_1 = pd.read_excel(excel_file, sheet_name='hoja1')
    resultado= df_1.loc[df_1['Rn'] >= Rnn]
    #resultado = df.loc[df['Rn'] >= Rnn ]
    #NPR
    if um<=420:
        

        NPR = resultado.iloc[0, resultado.columns.get_loc("NPR1")+0]

    if um>420 and um<=550:
        NPR = resultado.iloc[0, resultado.columns.get_loc("NPR2")+0]
    if um>550:
        NPR = resultado.iloc[0, resultado.columns.get_loc("NPR3")+0]

    #NPM
    if um<=145:
        NPM = resultado.iloc[0, resultado.columns.get_loc("NPM1")+0]
        
    if um>145 and um<=362:
        NPM = resultado.iloc[0, resultado.columns.get_loc("NPM1")+0]
    if um>362:
        NPM = resultado.iloc[0, resultado.columns.get_loc("NPM2")+0]

    #NPR = resultado.iloc[0, resultado.columns.get_loc("NPR")+1]
    #NPM = resultado.iloc[0, resultado.columns.get_loc("NPM")+2]

    ########################################## seleccionar el metodo de correccion de altura
    if h>1000:
        m=body["h_correction"]
        if m=="ansi":
            ka=(1/(1+1.25*10**(-4)*(h-1000)))
            NPR_NEW=NPR/ka
            NPM_NEW=NPM/ka
        if m=="iec":
            ka=(2.718281828**((h-1000)/8150))
            NPR_NEW=NPR*ka
            NPM_NEW=NPM*ka
    else:
        NPR_NEW=NPR
        NPM_NEW=NPM    
    ######################################### eleccion del bil
        
    BIL=1.25*NPR_NEW
    df_2 = pd.read_excel(excel_file, sheet_name='hoja2')
    B= df_2.loc[df_2['BIL'] >= BIL]
    BIL = B.iloc[0, B.columns.get_loc("BIL")+0]
    #BIL=math.ceil(BIL / 50) * 50

    k=1.15
    BSL=0.75*BIL
    BSL=int(BSL)
    while k>(BSL/NPM_NEW):
        BIL+=100
        BSL=0.75*BIL
    B2= df_2.loc[df_2['BIL'] >= BSL]
    BSL = B2.iloc[0, B2.columns.get_loc("BIL")+0]
    BSL=int(BSL)
    BIL=int(BIL)
    print("NPR elegido por el catalogo ABB sin correccion de altura, oxido de cinc: ",NPR,"kV")
    print("NPM elegido por el catalogo ABB sin correccion de altura, oxido de cinc: ",NPM,"kV")
    print("El BSL a escoger es de: ",BSL,"kV")

    print ("El BIL a escoger es de: ",BIL,"kV")
    ################################################ distancias de seguridad
    dminft=1.04*0.894**(-0.9)*(BIL/550)
    dminff=1.2*dminft
    VB=1.1*dminft
    print("Distancia minima fase tierra es: ",dminft,"m")
    print("Distancia minima fase fase es: ",dminff,"m")
    print("El valor basico de la subestacion es: ",VB,"m")
    ############################################# distancias de trabajo
    zona_cir=VB+2.25
    DH=VB+1.75
    DV=VB+1.25
    car=5.2
    hequi=2.3+0.0105*um
    hbarras=5+0.0125*um
    hrematelt=5+0.006*um
    print("La medida de la zona de circulacion es: ",zona_cir,"m")
    print("la  distancia horizontal es: ",DH,"m")
    print("La distancia vertical es: ",DV,"m")
    print("La distancia del galibo es de: ",car,"m")
    print("La altura de equpos es de: ",hequi,"m")
    print("La altura de barras colectoras es de: ",hbarras,"m")
    print("La altura de remate LT  es de: ",hrematelt,"m")
    result = {
        "npr": NPR,
        "npm": NPM,
        "bsl": BSL,
        "bil": BIL,
        "dminft": dminft,
        "dminff": dminff,
        "vb": VB,
        "zona_cir": zona_cir,
        "dh": DH,
        "dv": DV,
        "car": car,
        "hequi": hequi,
        "hbarras": hbarras,
        "hrematelt": hrematelt
    }
    return result
        
if __name__ == "__main__":
    print(lambda_handler({"body": {"vm": 500, "ke": 1.4, "h": 3500, "h_correction": "iec"}}, {}))

#print("el  ",BIL,BSL,NPR_NEW)
#print("el  ",NPM_NEW)
#print(df)
