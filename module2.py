import matplotlib
matplotlib.use('Agg')
from matplotlib import pyplot as plt
from matplotlib.colors import ListedColormap
import numpy as np
import pandas as pd
import os
import geopandas as gpd

# provide the path to the gdb and the layer name for the current year minus 1-3 and current year EMDS results from Keith's results

GDB_PATH = r"C:\Users\markhammond\Desktop\cflrp\TCA_CFLRP_Databases.gdb"

LAYER_CY_MINUS_3 = "TCA_Baseline_2020"

LAYER_CY_MINUS_2 = "TCA_Baseline_2021"

LAYER_CY_MINUS_1 = "TCA_Baseline_2022"

LAYER_CY = "TCA_Baseline_2023"

LAYERS = [LAYER_CY_MINUS_3, LAYER_CY_MINUS_2, LAYER_CY_MINUS_1, LAYER_CY]

YEAR_LIST = [2020, 2021, 2022, 2023]

OUTPUT_PATH = r"C:\Users\markhammond\Desktop\cflrp" # outpath for figures

# use this list once we clean up the field names
indicators = ["M_terrestrial_condition","M_disturbance_agents","M_biotic_agents",
                "M_tree_mortality","M_tree_mortality_05","M_tree_mortality_610",
                "M_abiotic_agents","M_uncharacteristic_disturbance","M_uncharacteristic_wildfire","M_moderate_severity",
               "M_low_severity",
               "M_road_density",
               "M_paved_roads","M_light_duty_roads","M_unimproved_roads","M_climate_exposure",
               "M_drought_impact","M_temperature_exposure","M_temp_exposure_spring",
               "M_temp_exposure_summer","M_temp_exposure_fall","M_temp_exposure_winter",
               "M_precipitation_exposure","M_precip_exposure_spring",
               "M_precip_exposure_spring_pct","M_precip_exposure_summer",
               "M_precip_exposure_summer_pct","M_precip_exposure_fall","M_precip_exposure_fall_pct",
               "M_precip_exposure_winter","M_precip_exposure_winter_pct",
               "M_air_pollution","M_nitrogen_deposition","M_ozone","M_ozone_accute_exposure","M_ozone_chronic_exposure",
               "M_vegetation_condition", "M_forest_condition",
               "M_vegetation_departure", "M_fire_deficit",
               "M_wildfire_hazard_potential",
               "M_wildfire_hazard_potential_HVH",
               "M_wildfire_hazard_potential_M",
               "M_insect_and_pathogen_hazard",
               "M_grassland_condition", "M_grassland_productivity", "M_grassland_encroachment"]
               

nameList = ["Terrestrial Condition", "Disturbance Agents and Stressors", "Biotic Agents",
                "Tree Mortality","Tree Mortality 0-5","Tree Mortality 6-10", 
                "Abiotic Agents", "Uncharacteristic Disturbance Events", "Uncharacteristic Wildfire", "Moderate Severity Fire", 
                "Low Severity Fire",
                "Road Density", 
                "Paved Roads","Light Duty Roads", "Unimproved Roads","Climate Exposure", 
                "Drought Impact", "Temperature Exposure", "Temp. Exposure Spring", 
                "Temp. Exposure Summer", "Temp. Exposure Fall", "Temp. Exposure Winter", 
                "Precipitation Exposure","Precip. Exposure Spring", 
                "Precip. Exposure Spring (Percent)", "Precip. Exposure Summer", 
                "Precip. Exposure Summer (Percent)", "Precip. Exposure Fall", "Precip. Exposure Fall (Percent)",
                "Precip. Exposure Winter", "Precip. Exposure Winter (Percent)", 
                "Air Pollution","Nitrogen Deposition", "Ozone", "Ozone Acute Exposure", "Ozone Chronic Exposure",
                "Vegetation Condition", "Forest Condition",
                "Vegetation Departure", "Fire Deficit",
                "Wildfire Hazard Potential", "Wildfire Hazard Potential (High/Very high)", "Wildfire Hazard Potential (Moderate)",
                "Insect and Disease Pathogen Hazard",
                "Grassland Condition", "Grassland Productivity", "Grassland Encroachment"]

nameDict = dict(zip(indicators, nameList))

# include only certain fields when reading in the data for testing and efficiency
include = ["tca_id","tca_acres", "cflrp_project_name", "M_terrestrial_condition"]

inds = ["M_terrestrial_condition"]

for var in inds:
    
    print(var)
    
    # initalize an empty list then
    lscapes = []
    acres = []

    for count, layer in enumerate(LAYERS):

        # print(count,layer)

        year = "_" + str(YEAR_LIST[count])
        # print(year)

        lyr = layer
        # print("reading {} feature class".format(lyr))
        gdf = gpd.read_file(GDB_PATH, driver = "FileGDB", layer = lyr, include_fields=include, where="cflrp_project_name='Southern Blues Restoration Coalition'") # subsets the data by row using the "where" query

        # convert gdf to df
        df = pd.DataFrame(gdf.drop(columns='geometry'))

        # reclassify the range of scores into the categories
        bins = [-1,-0.6,-0.2,0.2,0.6,1]
        labels=["Very Poor", "Poor", "Moderate", "Good", "Very Good"]
        df['Class'] = pd.cut(gdf[var], bins=bins, labels=labels, include_lowest=True)

        #get count by category
        count_by_class = df.groupby(['Class']).size()
        count_by_class.index.names = [var + '_class']
        
        lscapes.append(count_by_class)
        
        #get sum of acres
        acres_by_class = df.groupby(['Class'])['tca_acres'].sum()/1000
        acres_by_class.index.names = [var + '_acres']
        
        acres.append(acres_by_class)

    # %%

    # df_lscapes = pd.concat([lscapes[0], lscapes[1], lscapes[2],lscapes[3]], axis = 1).rename(columns = {0:"2020",1:"2021",2:"2022",3:"2023"})
    # print(df_lscapes)
    
    acres_df = pd.concat([acres[0], acres[1], acres[2], acres[3]], axis = 1).set_axis(['2020', '2021','2022','2023'], axis='columns')
    acres_df = acres_df.astype(int)
    print(acres_df)
    
    # %%
    

    # #plot the scores by category
    # fig, ax = plt.subplots(1, 1)
    # fig.set_size_inches(6,6)

    # # custom cmap
    # colors = ["#B61D36", "#FF9966", "#DFD696", "#00A3E6", "#003399"]
    # custom_cmap = ListedColormap(colors)

    # #get count by category
    # df.transpose().plot(kind='bar', stacked=True, ax=ax, colormap=custom_cmap)

    # plt.legend(loc='upper center', ncol=5, frameon=True, bbox_to_anchor=(0.5, 1.1), fancybox=True, shadow=True, fontsize=10)
    # ax.spines['right'].set_visible(False)
    # ax.spines['top'].set_visible(False)
    # plt.xlabel('Condition Class', size=14)
    # plt.xticks(rotation=45, ha = 'right', size=12)
    # plt.ylabel('Number of Landscapes', size=14)
    # plt.yticks(size=12)
    # plt.title(nameDict.get(var), size = 18, y=1.1)
    # plt.tight_layout()

    # print("saving condition/threat class figure for {}...".format(var))

    # outName = var + "_barchart_byClass"
    # outFile = os.path.join(OUTPUT_PATH, outName)
    # plt.savefig(outFile, dpi=300)

    # plt.close()

    # start table making

    HEADER_COLOR = '#00008B'

    
    #ROW_COLORS = [['#8b0038', 'w','w','w'],['#8b0038', 'w','w','w'],['#8b0038', 'w','w','w'],['#8b0038', 'w','w','w'],['#8b0038', 'w','w','w']]
    ROW_COLORS = ['#f1f1f2', 'w']
    #colors = ["#B61D36", "#FF9966", "#DFD696", "#00A3E6", "#003399"]

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

    # %%
    print(acres_df.values)

    # %%
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
        #if k ==(1,-1):
            #cell.set_facecolor("#B61D36")
        else:
            cell.set_facecolor(ROW_COLORS[k[0]%len(ROW_COLORS) ])
        if k ==(1,-1):
            cell.set_facecolor("#B61D36")
        if k ==(2,-1):
            cell.set_facecolor("#FF9966")
        if k ==(3,-1):
            cell.set_facecolor("#DFD696")
        if k ==(4,-1):
            cell.set_facecolor("#00A3E6")
        if k ==(5,-1):
            cell.set_facecolor("#003399")

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
    outName = var + "_table_byClass_acresttest"
    outFile = os.path.join(OUTPUT_PATH, outName)
    plt.savefig(outFile, dpi=300,bbox_inches=BBOX_INCHES)
    acresTable = outFile + ".png"
        
