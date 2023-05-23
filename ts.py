# -*- #################
# Descrição: CRIA LAYER DE TODAS AS CAMADAS DO DATAFRAME ATIVO PROXIMOS (1000M) DA LOCALIZAÇÃO SELECIONADA
# autor: Alberto Tuioshi Nagao - 10/04/2019

#     biblioteca Python para Excel 2010 "openpyxl" no pc atraves de:
#     no prompt "cmd": digite "cd C:\Python27\ArcGIS10.5\Scripts" 
#                      digite "pip install openpyxl" 
import openpyxl
import os, sys
import datetime
import arcpy  
from unicodedata import normalize


arcpy.env.overwriteOutput = True

def adicionalayerTOC(filel, outlayer, df):
    if type(filel) != unicode:
        a1 =filel
    else:
        a1 =normalize('NFKD', filel).encode('ASCII', 'ignore').replace(" ","")  # Retira os acentos e caracteres NÃO ASCII e os espaços
    arcpy.AddMessage("Adiciona layer TOC:{}".format(a1))
    arcpy.MakeFeatureLayer_management(a1, outlayer)     # Cria layer (temporário) 
    addlayer = arcpy.mapping.Layer(outlayer)
    desc = arcpy.Describe(filel)
    if desc.shapeType == "Polygon":
        arcpy.mapping.AddLayer(df, addlayer, "BOTTOM")  # Adiciona outlayer como último da lista do df
    else:
        arcpy.mapping.AddLayer(df, addlayer, "TOP")     # Adiciona outlayer como primeiro da lista do df
    
    del addlayer
    
def criarBuf(selec, df1, df, pasta):
    pontoAnalisado ="PontoArea"     
    arcpy.CopyFeatures_management(in_features=selec,    # 
        out_feature_class="AreaAnalisada",
        config_keyword="", 
        spatial_grid_1="0", 
        spatial_grid_2="0", 
        spatial_grid_3="0")

    desc = arcpy.Describe("AreaAnalisada")
    if desc.shapeType == "Polygon" :
        arcpy.FeatureToPoint_management(
            in_features="AreaAnalisada",                        
            out_feature_class="pontoDaArea", 
            point_location="CENTROID")
        pontoAnalisado = "pontoDaArea"
    elif desc.shapeType == "Polyline":
        arcpy.AddMessage("Rotina para linha a implementar: {}".format(selec))   
        sys.exit()  
    else:
        arcpy.AddMessage("Rotina para ponto a implementar: {}".format(selec))   
        sys.exit()  
        
    arcpy.Buffer_analysis(in_features=pontoAnalisado,   
        out_feature_class="RAIO1KM",
        buffer_distance_or_field=raio + " Meters", line_side="FULL", 
        line_end_type="ROUND", 
        dissolve_option="NONE", 
        dissolve_field="", method="PLANAR")
        
    layerTemplate = "AREAANALISADA.lyr"
    templayer = os.path.join(pasta, "Layer Templates")
    lyrFile = os.path.join(templayer, layerTemplate)    
    
    filel = "AreaAnalisada"          
    outlayer = "AREA ANALISADA"
    adicionalayerTOC(filel, outlayer, df)
    
    layerToUpd = arcpy.mapping.ListLayers(mxd, outlayer, df)[0] 
    arcpy.ApplySymbologyFromLayer_management(layerToUpd, lyrFile)
    
    if len(arcpy.mapping.ListLayoutElements(mxd, "LEGEND_ELEMENT")) > 0:
        legend = arcpy.mapping.ListLayoutElements(mxd, "LEGEND_ELEMENT")[0]
        legend.updateItem(layerToUpd, show_feature_count=False)
    
    if len(arcpy.mapping.ListDataFrames(mxd)) > 1:
        lyrFile = "//Layer Templates/AREAANALISADA_DF1.lyr"
        addlayer = arcpy.mapping.Layer(outlayer)
        arcpy.mapping.AddLayer(df1, addlayer, "TOP")
        layerToUpd = arcpy.mapping.ListLayers(mxd, outlayer, df1)[0]
        del addlayer    
        arcpy.ApplySymbologyFromLayer_management(layerToUpd, lyrFile)
        
    layerTemplate = "RAIO1KM.lyr"
    templayer = os.path.join(pasta, "Layer Templates")
    lyrFile = os.path.join(templayer, layerTemplate)    
    
    filel = "RAIO1KM"
    if raio != "1000":
        outlayer = "RAIO " + raio + "M"
    else:
        outlayer = "RAIO 1KM"
    adicionalayerTOC(filel, outlayer, df)
    
    layerToUpd = arcpy.mapping.ListLayers(mxd, outlayer, df)[0] 
    ext = layerToUpd.getExtent()                                # aplica escala na extensao deste layer "RAIO 1KM"
    df.extent = ext
    arcpy.ApplySymbologyFromLayer_management(layerToUpd, lyrFile)
    
    if len(arcpy.mapping.ListLayoutElements(mxd, "LEGEND_ELEMENT")) > 0:
        legend = arcpy.mapping.ListLayoutElements(mxd, "LEGEND_ELEMENT")[0]
        legend.updateItem(layerToUpd, show_feature_count=False)
        
    return pontoAnalisado

def layerOfSelection(inlayer, pontoAnalisado, df, pasta, raio):
    arcpy.SelectLayerByLocation_management(
        in_layer=inlayer.name,                      
        overlap_type="WITHIN_A_DISTANCE", 
        select_features=pontoAnalisado, 
        search_distance= raio + " Meters", 
        selection_type="NEW_SELECTION", 
        invert_spatial_relationship="NOT_INVERT")
            
    if type(inlayer.name) != unicode:
        a1 = inlayer.name + "_"
    else:
        a1 = normalize('NFKD', inlayer.name).encode('ASCII', 'ignore').replace(" ","") + "_"
    
    desc = arcpy.Describe(inlayer)
    geometry_type = desc.shapeType
    matchcount = int(arcpy.GetCount_management(inlayer.name)[0])  # conta qtos layers foram selecionados acima
    
    arcpy.AddMessage("Pesquisando:{}      cont:{}".format(a1, matchcount))  
    
    if matchcount == 0:
        arcpy.CreateFeatureclass_management("in_memory", a1, geometry_type, inlayer.name)
    else:
        arcpy.CopyFeatures_management(
            in_features=inlayer.name,                   
            out_feature_class=a1,
            config_keyword="", 
            spatial_grid_1="0", spatial_grid_2="0", spatial_grid_3="0")
        
    if "SAUDE" in inlayer.name.upper():
        outlayer = "SAUDE"
        layerTemplate = "SAUDE.lyr"
    elif "CEINF" in inlayer.name.upper():
        outlayer = "CEINF"
        layerTemplate = "CEINF.lyr"
    elif "CRECHE" in inlayer.name.upper():
        outlayer = "CRECHE"
        layerTemplate = "CRECHE.lyr"
    elif "MUNICIPA" in inlayer.name.upper() and ("ESCOLA" in inlayer.name.upper() or "ENSINO" in inlayer.name.upper()):
        outlayer = "ESCOLA MUNICIPAL"
        layerTemplate = "ESCOLAMUNICIPAL.lyr"
    elif "ESTADUA" in inlayer.name.upper() and ("ESCOLA" in inlayer.name.upper() or "ENSINO" in inlayer.name.upper()):
        outlayer = "ESCOLA ESTADUAL"
        layerTemplate = "ESCOLAESTADUAL.lyr"
    elif "ACADEMIA" in inlayer.name.upper():
        outlayer = "ACADEMIA AO AR LIVRE"
        layerTemplate = "ACADEMIA.lyr"
    elif "PROTECAO" in inlayer.name.upper() or "SOCIAL" in inlayer.name.upper():
        outlayer = "C R A S"
        layerTemplate = "CRAS.lyr"
    elif "REA" in inlayer.name.upper() and "BLIC" in inlayer.name.upper():
        outlayer = "AREA PUBLICA"
        layerTemplate = "AREAPUBLICA.lyr"
    elif "RREGO" in inlayer.name.upper():
        outlayer = "CORREGO"
        layerTemplate = "CORREGO.lyr"
    elif "IMOVE" in inlayer.name.upper():
        outlayer = "IMOVEL"
        layerTemplate = "IMOVEL.lyr"
    else:
        outlayer = inlayer.name.upper()
        layerTemplate = os.path.splitext(a1)[0] + ".lyr"
            
    adicionalayerTOC(a1, outlayer, df)
    
    layerToUpd = arcpy.mapping.ListLayers(mxd, outlayer, df)[0]
    templayer = os.path.join(pasta, "Layer Templates")
    lyrFile = os.path.join(templayer, layerTemplate)    
    
    if os.path.isfile(lyrFile):
        arcpy.ApplySymbologyFromLayer_management(layerToUpd, lyrFile)
        if outlayer == "CORREGO" or outlayer == "AREA PUBLICA":         
            if len(arcpy.mapping.ListLayoutElements(mxd, "LEGEND_ELEMENT")) > 0:
                legend = arcpy.mapping.ListLayoutElements(mxd, "LEGEND_ELEMENT")[0]
                legend.updateItem(layerToUpd, show_feature_count=False)         #Retira a Contagem autom. de features na legenda
        else:
            legend = arcpy.mapping.ListLayoutElements(mxd, "LEGEND_ELEMENT")[0]
            legend.updateItem(layerToUpd, show_feature_count=True)
        arcpy.mapping.RemoveLayer(df, inlayer)
        
if __name__ == '__main__':  
    selec = arcpy.GetParameterAsText(0)
    processo = arcpy.GetParameterAsText(1)
    requerente = arcpy.GetParameterAsText(2)
    planilha = arcpy.GetParameterAsText(3)
    raio = arcpy.GetParameterAsText(4)
    datadia = arcpy.GetParameterAsText(5)


    #pasta = arcpy.Describe(selec).path                                         # caminho até pasta do selec
    pasta = os.path.dirname(os.path.realpath(sys.argv[0]))                      # caminho até pasta do script
    parcelamen = "PARCELAMEN"
    tipolog = "TIPOLOG"
    ruaimo = "RUAIMO"
    nrporta = "NRPORTA" 
    
    mxd = arcpy.mapping.MapDocument("CURRENT")  
    df = arcpy.mapping.ListDataFrames(mxd)[0]
    df1 = df
    if len(arcpy.mapping.ListDataFrames(mxd)) > 1:
        df1 = arcpy.mapping.ListDataFrames(mxd)[1]
    layers = arcpy.mapping.ListLayers(df)
    tt = [arcpy.Describe(lyr.name).FIDset.split("; ") for lyr in layers if lyr.name == selec]   # Lista de "features" selecionados
    if len(tt[0]) == 1 and tt[0] == [u'']:
        arcpy.AddMessage("* * *  NENHUM FEATURE SELECIONADO     * * *")     
        sys.exit()  
    elif len(tt[0]) > 1:
        arcpy.Dissolve_management(in_features=selec, 
            out_feature_class=selec + "_Dissolve", 
            dissolve_field="", 
            statistics_fields=[["PARCELAMEN","FIRST"], ["TIPOLOG","FIRST"], ["RUAIMO","FIRST"], ["NRPORTA","FIRST"]],
            multi_part="MULTI_PART", 
            unsplit_lines="DISSOLVE_LINES")
        selec = selec + "_Dissolve"
        parcelamen = "FIRST_PARCELAMEN"
        tipolog = "FIRST_TIPOLOG"
        ruaimo = "FIRST_RUAIMO"
        nrporta = "FIRST_NRPORTA"
        arcpy.AddMessage("* * *                                 * * *")     
        arcpy.AddMessage("* * *  MAIS DE UM FEATURE SELECIONADO * * *") 
        arcpy.AddMessage("* * *                                 * * *")         
        arcpy.AddMessage("* * *      EXECUTANDO DISSOLVE        * * *")     
        arcpy.AddMessage("* * *                                 * * *")     
    
    ptoAnalisado = criarBuf(selec, df1, df, pasta)
    
    for lyr in layers:
        if lyr.isFeatureLayer:
            layerOfSelection(lyr, ptoAnalisado, df, pasta, raio)
# calcula area total de territorial e particular:
    arcpy.Select_analysis(in_features="IMOVEL", 
        out_feature_class="IMOVEL__Select", 
        where_clause="USOIMOVEL = 'TERRITORIAL' AND PATRIMONIO = 'PARTICULAR'")
    total_area = sum(int(i[0]) for i in arcpy.da.FeatureClassToNumPyArray("IMOVEL__Select", 'AREALOTE'))    
# localiza o bairro:
    x = datetime.datetime.now()
    mxd.activeView ="New Data Frame"
    arcpy.SpatialJoin_analysis(target_features="AREA ANALISADA", 
        join_features="Bairros_region", 
        out_feature_class="AreaAnalisada_SpatialJoin_" + x.strftime("%M%S"), 
        join_operation="JOIN_ONE_TO_ONE", 
        join_type="KEEP_COMMON", 
        field_mapping='INSCANT "INSCANT" true true false 11 Text 0 0 ,First,#,AREA ANALISADA,INSCANT,-1,-1;NOME "NOME" true true false 40 Text 0 0 ,First,#,Bairros_region,NOME,-1,-1;POP2010 "POP2010" true true false 31 Double 15 30 ,First,#,Bairros_region,POP2010,-1,-1;DENS2010 "DENS2010" true true false 8 Float 2 7 ,First,#,Bairros_region,DENS2010,-1,-1', 
        match_option="INTERSECT", 
        search_radius="", 
        distance_field_name="")
    bairro = arcpy.da.SearchCursor("AreaAnalisada_SpatialJoin_"+ x.strftime("%M%S"), ("NOME",)).next()[0]           
    parcelamento = arcpy.da.SearchCursor("AREA ANALISADA", (parcelamen,)).next()[0]  
    tipo, rua, nr = arcpy.da.SearchCursor("AREA ANALISADA", (tipolog, ruaimo, nrporta)).next()

    arcpy.management.AddField("AREA ANALISADA", "requere", "TEXT")
    arcpy.management.AddField("AREA ANALISADA", "proces", "TEXT")
    arcpy.management.AddField("AREA ANALISADA", "data", "TEXT")
 
    arcpy.AddMessage("* * *  Colunas criadas na tabela     * * *") 
    with arcpy.da.UpdateCursor("AREA ANALISADA", "requere") as cursor:
        for row in cursor:
            row[0] = requerente
            cursor.updateRow(row)
    with arcpy.da.UpdateCursor("AREA ANALISADA", "proces") as cursor:
        for row in cursor:
            row[0] = processo
            cursor.updateRow(row)
    with arcpy.da.UpdateCursor("AREA ANALISADA", "data") as cursor:
        for row in cursor:
            row[0] = datadia
            cursor.updateRow(row)

    arcpy.AddMessage("* * *  add info na tabela     * * *") 
        
    arcpy.AddMessage("* {} {} {} *".format(tipo, rua, nr.lstrip("0")))
    if type(bairro) == unicode:
        arcpy.AddMessage("* BAIRRO:{} PARCELAMENTO:{} *".format(bairro.encode('utf-8'), parcelamento))
    
    local = tipo.capitalize() + " " + rua.title() + ", " + nr.lstrip("0")
    mxd.activeView ="Layers"    
# ler excel com densidade demográfica
    book = openpyxl.load_workbook(planilha)
    sheet = book.active
    cells = sheet['B3': 'C82']
    b1 = normalize('NFKD', bairro).encode('ASCII', 'ignore').replace(" ","")
    densid = 0
    for b, d in cells:  
        b2 = normalize('NFKD', b.value).encode('ASCII', 'ignore').replace(" ","")
        if b2 == b1:
            if type(d.value) is str:
                densid = float(d.value.replace(",",".").replace(" ",""))
            elif type(d.value) is float:
                densid = d.value
            else:
                print('type(d.value) is {}'.format(type(d.value)))
                densid = float(str(d.value).replace(",","."))
            break       
    if densid == 0:
        arcpy.AddMessage("* * *  NENHUM BAIRRO CORRESPONDENTE NA PLANILHA DE DENSIDADE DEMOG.  * * *")      
        sys.exit()  
# altera text element no layout:
    hec = total_area / 1000000
    prevPop = int(round(hec * densid))
    arcpy.AddMessage("{} = total_area:{}     *    dens:{}".format(prevPop, total_area, densid))
    for TextElement in arcpy.mapping.ListLayoutElements(mxd, "TEXT_ELEMENT"):
        a1 =normalize('NFKD', TextElement.text).encode('ASCII', 'ignore').replace(" ","")
        if "Previs" in TextElement.text:
            TextElement.text = TextElement.text.replace("NNNNN", str(int(round(total_area/1000000*densid))))
        elif "Proc" in TextElement.text:
            TextElement.text = TextElement.text.replace("NNNNN", processo).replace("RRRRR", requerente.upper()).replace("LLLLL", local).replace("BBBBB", bairro.upper()).replace("PPPPP", parcelamento.title())
        
    arcpy.RefreshTOC()
    arcpy.RefreshActiveView() 
          
    NewProj = requerente.upper()+ ".mxd"
    mxd.saveACopy(os.path.join(pasta, NewProj))
    newmxd = os.path.join(pasta, NewProj)
    #open new mxd 
    os.startfile(os.path.join(pasta, NewProj))
    #closing old project
    from signal import SIGTERM
    os.kill(os.getpid(),SIGTERM)
          
          
        





















