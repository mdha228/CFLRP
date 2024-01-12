from docx import Document
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
from docx2pdf import convert
import os
import datetime
import calendar
import matplotlib
matplotlib.use('Agg')
from matplotlib import pyplot as plt
from matplotlib.colors import ListedColormap
import numpy as np
import pandas as pd
import os
import arcpy

def updateTitleDoc(projectArea, projectsDict, pathToTitles, docxPath):
    ##update forest name, region, and date in word doc and convert to pdf to create first 2 pages of TCA report
    outputName = os.path.join(pathToTitles,'CFLRP_TitlePages_{}.docx'.format(projectArea))
    #if os.path.exists(outputName) == overwrite or os.path.exists(outputName) == False: 
    document = Document(docxPath)
    region = 'Region '+str(projectsDict[projectArea])
    internal = document.paragraphs[2]
    txt = internal.text[:5]
    internal.text = ''
    obj_styles = document.styles
    obj_charstyle = obj_styles.add_style('CommentsStyle3', WD_STYLE_TYPE.CHARACTER)
    obj_font = obj_charstyle.font
    obj_font.size = Pt(12)
    obj_font.name = 'Arial'
    date = calendar.month_name[datetime.datetime.now().month] + ' '+str(datetime.datetime.now().year)
    internal.add_run(date, style = 'CommentsStyle3')
    section = document.paragraphs[5]       
    fontsize = 30 
    section.text = ''
    obj_styles = document.styles
    obj_charstyle = obj_styles.add_style('CommentsStyle', WD_STYLE_TYPE.CHARACTER)
    obj_font = obj_charstyle.font
    obj_font.size = Pt(fontsize)
    obj_font.name = 'Arial'
    section.add_run(projectArea, style = 'CommentsStyle')
    section2 = document.paragraphs[6]
    fontsize = 12
    section2.text = ''
    obj_charstyle2 = obj_styles.add_style('CommentsStyle2', WD_STYLE_TYPE.CHARACTER)
    obj_font2 = obj_charstyle2.font
    obj_font2.italic = True
    obj_font2.size = Pt(fontsize)
    obj_font2.name = 'Arial'
    section2.add_run(region, style = 'CommentsStyle2')
    if os.path.exists(outputName):
        os.remove(outputName)
    document.save(outputName)
    pdfOutputLoc = outputName[0:-4] + "pdf"
    print(f"Output location for pdf: {pdfOutputLoc}")
    print(f"Output location for word doc: {outputName}\n")
    print ('*** Do not open any docx files until all have been converted (this will overwrite current pdfs) ***\n')
    convert(outputName, pdfOutputLoc)

def feature_class_to_pandas_data_frame(feature_class, field_list):
    """
    Load data into a Pandas Data Frame for subsequent analysis.
    :param feature_class: Input ArcGIS Feature Class.
    :param field_list: Fields for input.
    :return: Pandas DataFrame object.
    """
    return pd.DataFrame(
        arcpy.da.FeatureClassToNumPyArray(
            in_table=feature_class,
            field_names=field_list,
            skip_nulls=False,
            null_value=-99999
        )
    )
def exportAndAppendToFinalPDF(srcLayout, tempPath, finalPath):
        srcLayout.exportToPDF(tempPath)#,resolution=500, output_as_image=True) #output as image to convert vector parts of pdf to raster (helps with width issue of LTA outline)
        finalPath.appendPages(tempPath)
GDB_PATH = r"C:\Users\markhammond\Desktop\cflrp\TCA_CFLRP_Databases.gdb"

arcpy.env.workspace = GDB_PATH
fcs = arcpy.ListFeatureClasses()
fc_path_list = [ ]
for fc in fcs:
    fc_path = os.path.join(GDB_PATH , fc)
    fc_path_list.append(fc_path)

LAYERS = ["TCA_Baseline_2020", "TCA_Baseline_2021", "TCA_Baseline_2022", "TCA_Baseline_2023"]

YEAR_LIST = [2020, 2021, 2022, 2023]

OUTPUT_PATH = r"C:\Users\markhammond\Desktop\cflrp\output" # outpath for figures

ANCILLARY_FILES_PATH = r"C:\Users\markhammond\Desktop\cflrp\ancillary"

APRX = arcpy.mp.ArcGISProject(r"C:\Users\markhammond\Desktop\cflrp\CFLRP_SouthernBlues.aprx")

SYMBOLOGY_LAYER_PATH = os.path.join(ANCILLARY_FILES_PATH,'sym.lyrx')

TITLE_PAGES_TEMPLATE = os.path.join(ANCILLARY_FILES_PATH, 'TITLE_PAGE_TEMPLATE.docx')

FINAL_OUTPUT_PATH = os.path.join(OUTPUT_PATH, "FinalOutput")

### path for all generated pages per area.
OUTPUT_TITLE_PAGES_PATH = os.path.join(OUTPUT_PATH, "TitlePages")

APPENDIX_PDF_PATH = os.path.join(ANCILLARY_FILES_PATH, "CONUS_Appendix_Table.pdf")

GLOSSARY_PDF_PATH = os.path.join(ANCILLARY_FILES_PATH, "TCA_Glossary_CONUS.pdf")

projectsDict = {"Southern Blues Restoration Coalition": "6"}
cflrp_project_areas = list(projectsDict.keys())
layout_text = {'M_tree_mortality': ['Tree Mortality', 'Significant tree mortality caused by insects and pathogen outbreaks','Forested Systems'], 'M_fire_deficit':['Ecological Process Departure','Fire deficit due to fire suppression','All Systems']}

# include only certain fields when reading in the data for testing and efficiency
include = ["tca_id","tca_acres", "cflrp_project_name", "M_fire_deficit", "M_tree_mortality"]

inds = ["M_fire_deficit","M_tree_mortality"]
 
for area in cflrp_project_areas:
    PDF_path = os.path.join(FINAL_OUTPUT_PATH,"CFLRP_{}_Summary.pdf".format(area)) 
    finalPDF = arcpy.mp.PDFDocumentCreate(PDF_path)
    titlePage = os.path.join(OUTPUT_TITLE_PAGES_PATH, 'CFLRP_TitlePages_'+area+".pdf")
    updateTitleDoc(area, projectsDict, OUTPUT_TITLE_PAGES_PATH, TITLE_PAGES_TEMPLATE)
    finalPDF.appendPages(titlePage)
    
    for ind in inds:
        # initalize an empty lists
        acres_results = []
        count_results = []

        for fc in fc_path_list:
            query_project = 'cflrp_project_name' + " = " + "'"+area+"'"   
            year = fc[-4:]
            forests_subset = GDB_PATH + "_subset"
            queriedLayer = arcpy.management.MakeFeatureLayer(fc, forests_subset, query_project).getOutput(0)
            symbologyFields = "VALUE_FIELD {} {}".format(ind,ind)
            arcpy.management.ApplySymbologyFromLayer(queriedLayer, SYMBOLOGY_LAYER_PATH, symbologyFields)
            current_map = APRX.listMaps(year)[0]
            current_map.addLayer(queriedLayer)
            df = feature_class_to_pandas_data_frame(queriedLayer, include)
            bins = [-1,-0.6,-0.2,0.2,0.6,1]
            labels=["Very Poor", "Poor", "Moderate", "Good", "Very Good"]
            df['Class'] = pd.cut(df[ind], bins=bins, labels=labels, include_lowest=True)

            #get count by category
            count_by_class = df.groupby(['Class']).size()
            count_by_class.index.names = [ind + '_class']
            print(count_by_class)
            #get acres by category
            acres_by_class = df.groupby(['Class'])[include[1]].sum()/1000
            acres = acres_by_class.astype(int)
            acres.index.names = [ind + '_class']
            print(acres)

            acres_results.append(acres)
            count_results.append(count_by_class)
        count_df = pd.concat([count_results[0], count_results[1], count_results[2],count_results[3]], axis = 1).rename(columns = {0:"2020",1:"2021",2:"2022",3:"2023"})
        print(count_df)
        acres_df = pd.concat([acres_results[0], acres_results[1], acres_results[2],acres_results[3]], axis = 1).rename(columns = {0:"2020",1:"2021",2:"2022",3:"2023"})
        acres_df.columns = ['2020', '2021', '2022', '2023']
        print(acres_df)
        
        #create bar chart
        fig, ax = plt.subplots(1, 1)
        fig.set_size_inches(10,10)
        # custom cmap
        colors = ["#B61D36", "#FF9966", "#DFD696", "#00A3E6", "#003399"]
        custom_cmap = ListedColormap(colors)
        #get count by category
        count_df.transpose().plot(kind='barh', stacked=True, ax=ax, colormap=custom_cmap)
        plt.legend(loc='upper center', ncol=5, frameon=True, bbox_to_anchor=(0.5, 1.1), fancybox=True, shadow=True)
        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        plt.xlabel('Condition Class')
        plt.xticks(rotation=45, ha = 'right')
        plt.ylabel('Number of Landscapes')
        #plt.title(nameDict.get(var), size = 22, y=1.1)
        plt.tight_layout()
        print("saving condition/threat class figure for {}...".format(ind))
        outName = ind + "_barchart_byClass"
        outFile = os.path.join(OUTPUT_PATH, outName)
        plt.savefig(outFile, dpi=300)
        barChart = outFile + ".png"
        plt.close()
        
        #table variables
        HEADER_COLOR = '#00008B'

        ROW_COLORS = ['#f1f1f2', 'w']

        EDGE_COLOR='w'

        BBOX = [0, 0, 1, 1]

        BBOX_INCHES = 'tight'

        FONT_SIZE = 18

        HEADER_COLUMNS = 0

        MPL_COL_WIDTH = 1.3

        MPL_ROW_HEIGHT = 0.5

        MPL_FINAL_COL_WIDTH = 1.9

        MPL_FINAL_ROW_HEIGHT = 0.7

        COLOR_ARR = np.array([
        [0,0,139],
        [115,178,255],
        [245,202,122],
        [255,167,127],
        [255,0,0],
        [211,211,211]])
        
        #create number of landscapes table
        heading = "Number of Landscapes"
        size = (np.array(count_df.shape[::-1]) + np.array([MPL_COL_WIDTH, 1])) * np.array([MPL_COL_WIDTH, MPL_ROW_HEIGHT])
        fig, ax = plt.subplots(figsize=size)
        ax.axis('off')
        mpl_table_header = ax.table(cellText=[['']],colLabels = [heading],loc='top')
        mpl_table = ax.table(cellText=count_df.values, bbox=BBOX, colLabels=count_df.columns,rowLabels= count_df.index,loc='top')
        #ax.axhline(y=1, color=header_color, lw=1,alpha=1)  #alpha=1, 
        ax.axvline(x=0,color='w', alpha=1, lw=1)
        ax.axvline(x=0,color='w', alpha=1, lw=1)
        mpl_table.auto_set_font_size(False)
        mpl_table.set_fontsize(FONT_SIZE)
        mpl_table_header.auto_set_font_size(False)
        mpl_table_header.set_fontsize(FONT_SIZE)
        for k, cell in mpl_table._cells.items():
            cell.set_edgecolor(EDGE_COLOR)
            if k[0] == 0 or k[1] < HEADER_COLUMNS:
                cell.set_text_props(weight='bold', color='w')
                cell.set_facecolor(HEADER_COLOR)

                if k[0] == 0:
                    cell.set_text_props(weight='bold', color='w')#, va ='top')  
            else:
                cell.set_facecolor(ROW_COLORS[k[0]%len(ROW_COLORS) ])
        for k, cell in mpl_table_header._cells.items():
            cell.set_edgecolor(EDGE_COLOR)
            if k[0] == 0 or k[1] < HEADER_COLUMNS:           
                cell.set_facecolor(HEADER_COLOR)
                #h = cell.get_height()
                #cell.set_height(h/1.5)
                cell.set_text_props(weight='bold', color='w')#, verticalalignment ='bottom')
        #adjust height of first 'row' of initial table
        cellDict = mpl_table_header.get_celld()
        cellDict[(1,0)].set_height(0)
        if len(heading) > 25:
            cellDict[(0,0)].set_height(1/3.8)#6.) 
        else:
            cellDict[(0,0)].set_height(1/6.) 
        outName = ind + "_table_byClass"
        outFile = os.path.join(OUTPUT_PATH, outName)
        plt.savefig(outFile, dpi=300,bbox_inches=BBOX_INCHES)
        countTable = outFile + ".png"
        
        #create acres table
        heading = "Area (in thoudsands of acres)"
        size = (np.array(acres_df.shape[::-1]) + np.array([MPL_COL_WIDTH, 1])) * np.array([MPL_COL_WIDTH, MPL_ROW_HEIGHT])
        fig, ax = plt.subplots(figsize=size)
        ax.axis('off')
        mpl_table_header = ax.table(cellText=[['']],colLabels = [heading],loc='top')
        mpl_table = ax.table(cellText=acres_df.values, bbox=BBOX, colLabels=acres_df.columns,rowLabels= acres_df.index,loc='top')
        #ax.axhline(y=1, color=header_color, lw=1,alpha=1)  #alpha=1, 
        ax.axvline(x=0,color='w', alpha=1, lw=1)
        ax.axvline(x=0,color='w', alpha=1, lw=1)
        mpl_table.auto_set_font_size(False)
        mpl_table.set_fontsize(FONT_SIZE)
        mpl_table_header.auto_set_font_size(False)
        mpl_table_header.set_fontsize(FONT_SIZE)
        for k, cell in mpl_table._cells.items():
            cell.set_edgecolor(EDGE_COLOR)
            if k[0] == 0 or k[1] < HEADER_COLUMNS:
                cell.set_text_props(weight='bold', color='w')
                cell.set_facecolor(HEADER_COLOR)

                if k[0] == 0:
                    cell.set_text_props(weight='bold', color='w')#, va ='top')  
            else:
                cell.set_facecolor(ROW_COLORS[k[0]%len(ROW_COLORS) ])
        for k, cell in mpl_table_header._cells.items():
            cell.set_edgecolor(EDGE_COLOR)
            if k[0] == 0 or k[1] < HEADER_COLUMNS:           
                cell.set_facecolor(HEADER_COLOR)
                #h = cell.get_height()
                #cell.set_height(h/1.5)
                cell.set_text_props(weight='bold', color='w')#, verticalalignment ='bottom')
        #adjust height of first 'row' of initial table
        cellDict = mpl_table_header.get_celld()
        cellDict[(1,0)].set_height(0)
        if len(heading) > 25:
            cellDict[(0,0)].set_height(1/3.8)#6.) 
        else:
            cellDict[(0,0)].set_height(1/6.) 
        outName = ind + "_table_byClass_acres"
        outFile = os.path.join(OUTPUT_PATH, outName)
        plt.savefig(outFile, dpi=300,bbox_inches=BBOX_INCHES)
        acresTable = outFile + ".png"
        
        #update layout
        lay = APRX.listLayouts("Indicator")[0]
        layout_text = {'M_tree_mortality': ['Tree Mortality', 'Significant tree mortality caused by insects and pathogen outbreaks','Forested Systems'], 'M_fire_deficit':['Ecological Process Departure','Fire deficit due to fire suppression','All Systems']}
        text = lay.listElements("text_element")
        pictures = lay.listElements("picture_element")
        for elm in text:
            if elm.name =='Text 7':
                elm.text = cflrp_project_areas[0]
            elif elm.name =='Text':
                elm.text = layout_text[ind][0]
            elif elm.name =='Text 1':
                elm.text = layout_text[ind][1]
            elif elm.name =='Text 2':
                elm.text = layout_text[ind][2]
        for pic in pictures:
            if pic.name =='Picture 2':
                pic.sourceImage = barChart
            elif pic.name =='Picture 1':
                pic.sourceImage = countTable
            elif pic.name =='Picture 3':
                pic.sourceImage = acresTable
        #TEMP_PDF_PATH = os.path.join(OUTPUT_PATH, '{}_temp.pdf'.format(ind))
        exportAndAppendToFinalPDF(lay, TEMP_PDF_PATH, finalPDF)
    finalPDF.appendPages(APPENDIX_PDF_PATH)
    finalPDF.appendPages(GLOSSARY_PDF_PATH)
    finalPDF.saveAndClose()
