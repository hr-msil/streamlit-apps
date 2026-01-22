import numpy as np
import pandas as pd
import streamlit as st


#--------- LECTURA de archivos ---------------------

def lectura_archivo_prod(archivo_prod:str):
    '''
    Convierte en DataFrame el excel subido del sistema. Ordena los legajos por orden numérico.
    
    :param archivo_prod: Nombre del archivo .xlsx
    :type archivo_prod: Str
    :return DataFrame. 
    '''

    df_prod = pd.read_excel(archivo_prod)
    df_prod.sort_values(by = "Legajo") #Ordeno según legajo

    return df_prod

def lectura_archivo_dec(archivo_dec: str):
    '''
    Convierte en DataFrame los archivos subidos con extensión .csv. Si al convertirlo tiene más de 4 columnas, se toman las primeras cuatro.
    
    :param archivo_dec: String. Nombre del archivo .csv
    :type archivo_dec: Str.
    :return DataFrame.
    '''

    df_decreto = pd.read_csv(archivo_dec,header=None)

    if df_decreto.shape[1] > 4:

        df_decreto = df_decreto.iloc[:,:4]

    df_decreto.columns = ["Legajo", "Nula", "Nula2", "Importe"] #Renombro las columnas
    df_decreto = df_decreto.dropna(how="all") # Elimino las filas con algún Nan
    df_decreto["Legajo"] = df_decreto["Legajo"].astype('Int64') #Cambio tipo del Legajo para que tipe con df_prod
    df_decreto.sort_values(by = "Legajo") #Ordeno según legajo

    return df_decreto

#--------- LIMPIEZA DE LA COLUMNA DECRETOS -----------

def limpieza_decreto(df: pd.DataFrame) -> None:
    '''
    Modifica la columna "Leyenda" del dataFrame correspondiente al archivo subido del sistema para que el decreto quede de la forma num_dec/año.
    
    :param df: Description
    :type df: pd.DataFrame
    '''

    cant_prod = df.shape[0]

    for i in range(cant_prod):

        leyenda = df.iloc[i]["Leyenda"]
        if pd.isna(leyenda):
            # + 1 por el index de python, 1 por el encabezado de Excel
            st.write("La fila ", i + 2, " no tiene leyenda detallada." )
            st.divider()
        else:
            decreto_prod = leyenda.split(" ")[1] #VER Qué pasa si no se puede splitear por espacio?? tira error!
            df.loc[i,"Leyenda"] = decreto_prod

    df["Leyenda"] = df["Leyenda"].astype('str')



#--------- OBTENGO nombre del decreto, según nombre del archivo-----

def obtener_decreto(nombre_archivo: str) -> str:
    '''
    Dado el nombre del archivo, le quita la extensión .csv. Si el decreto está separado por "-", lo convierte a la forma nro_dec/año.
    
    :param nombre_archivo: Description
    :type nombre_archivo: str
    :return: Devuelve  el nombre del archivo sin la extensión.
    :rtype: str
    '''

    decreto = nombre_archivo.split(".")[0] #Con esto saco la extension .csv
    decreto = decreto.split(" ")
    decreto = decreto[0].split("-")[0] + "/" + decreto[0].split("-")[1] #Lo renombro a tipo nro_dec/año

    return decreto




#--------- FUNCION PRINCIPAL -------------------------

def comparar(df_prod_dec: pd.DataFrame, df_dec: pd.DataFrame, nombre_original: str) -> None:
    '''
    Toma el df de productividades filtrado por decreto y se fija si encuentra o no el monto correspondiente a cada legajo

    :param df_prod_dec: DataFrame de productividades filtrado por decreto.
    :type df_prod_dec: pd.DataFrame
    :param df_dec: DataFrame del decreto correspondiente
    :type df_dec: pd.DataFrame
    :param nombre_original: Description
    :type nombre_original: str
    '''

    cant_prod_dec = df_prod_dec.shape[0]
    cant_dec = df_dec.shape[0]
    #decreto = df_prod_dec["Leyenda"].unique()[0] #Nombre del decreto

    for i in range(cant_prod_dec):

        legajo = df_prod_dec.iloc[i]["Legajo"]
        importe = df_prod_dec.iloc[i]["Importe"]

        #Busco si existe la fila en el dataFrame correspondiente al decreto

        existe_en_csv = False

        for j in range(cant_dec):

            legajo_dec = df_dec.iloc[j]["Legajo"]
            importe_dec = df_dec.iloc[j]["Importe"]

            if importe_dec == importe and legajo_dec == legajo:
                
                existe_en_csv = True

        if existe_en_csv == False: #Agregar a un dataFrame global que sea el de diferencias

            legajos.append(legajo)
            importes.append(importe)

            



#--------- STREAMLIT -------------------------------

st.title("📝 PRODUCTIVIDADES")

st.divider()

tab1,tab2 = st.tabs(["Subir archivos", "Ver resultados"])

with tab1:

    st.markdown("Subir los archivos de productividades correspondientes a lo arrojado por el sistema")

    archivos_prod = st.file_uploader("Seleccionar archivo", type = "xlsx",key = "productividades",accept_multiple_files=True)

    st.markdown("Subir los archivos .csv que se quieren comparar")

    archivos_dec = st.file_uploader("Seleccionar archivo", type = "csv", key = "decreto",accept_multiple_files=True)


#-------- LECTURA Y LIMPIEZA de los archivos --------

with tab2:

    decretos_originales = []
    decretos = []

    if archivos_prod and archivos_dec:

        dfs_prod = []

        for archivo_prod in archivos_prod:

            df_prod = lectura_archivo_prod(archivo_prod)
            limpieza_decreto(df_prod)
            dfs_prod.append(df_prod)

        df_productividades = pd.concat(dfs_prod,ignore_index=True)
        
        sin_diferencias = []
        dfs_inconsistencias = []
        nombres_inconsistencias = []

        for archivo in archivos_dec:

            #--------- CREO df_diferencias ------------------------
            # Acá vamos a guardar todas las productividades correspondientes al archivo que se carga del SISTEMA
            # que no se encuentren en el csv correspondiente al decreto
            #hago un archivo de diferencias por decreto

            legajos = []
            importes = []

            nombre_original = archivo.name.split(".")[0]
            decreto = obtener_decreto(archivo.name)
            
            df_dec = lectura_archivo_dec(archivo)

            df_prod_dec = df_productividades[df_productividades["Leyenda"] == decreto]

            comparar(df_dec,df_prod_dec,nombre_original)

            df_diferencias = pd.DataFrame({"Legajo": legajos, "Importe": importes})

            if df_diferencias.shape[0] == 0:

                sin_diferencias.append(nombre_original)

            else:

                dfs_inconsistencias.append(df_diferencias)
                nombres_inconsistencias.append(nombre_original)

        if len(sin_diferencias) != 0:

            st.write("No se encontraron inconsistencias en los siguientes archivos: ")
            
            for no_diferencia in sin_diferencias: 

                st.write(""" - """, no_diferencia)
            
        tabs = st.tabs(nombres_inconsistencias)

        for i, df in enumerate(dfs_inconsistencias):

            with tabs[i]:

                st.write("En el archivo: ",nombres_inconsistencias[i]," no se encontraron estos importes para estos legajos:  ")

                st.write(df)









