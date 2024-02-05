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

OUTPUT_PATH = r"C:\Users\markhammond\Desktop\cflrp"

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
include = ["tca_id","tca_acres", "cflrp_project_name","M_terrestrial_condition"]

inds = ["M_terrestrial_condition"]

for var in inds:
    
    print(var)
    
    # initalize an empty list then
    results = []

    for count, layer in enumerate(LAYERS):

        print(count,layer)

        year = "_" + str(YEAR_LIST[count])
        print(year)

        lyr = layer
        print("reading {} feature class".format(lyr))
        gdf = gpd.read_file(GDB_PATH, driver = "FileGDB", layer = lyr, include_fields=include) # subsets the data by row using the "where" query

        # convert gdf to df
        df = pd.DataFrame(gdf.drop(columns='geometry'))

        # reclassify the range of scores into the categories
        bins = [-1,-0.6,-0.2,0.2,0.6,1]
        labels=["Very Poor", "Poor", "Moderate", "Good", "Very Good"]
        df['Class'] = pd.cut(gdf[var], bins=bins, labels=labels, include_lowest=True)

        #get count by category
        count_by_class = df.groupby(['Class']).size()
        count_by_class.index.names = [var + '_class']
        
        results.append(count_by_class)



    # %%

    df = pd.concat([results[0], results[1], results[2],results[3]], axis = 1).rename(columns = {0:"2020",1:"2021",2:"2022",3:"2023"})
    print(df)

    # %%
    

    # #plot the scores by category
    fig, ax = plt.subplots(1, 1)
    fig.set_size_inches(6,6)

    # # custom cmap
    colors = ["#B61D36", "#FF9966", "#DFD696", "#00A3E6", "#003399"]
    custom_cmap = ListedColormap(colors)

    # #get count by category
    df.transpose().plot(kind='bar', stacked=True, ax=ax, colormap=custom_cmap, legend=False, width = 0.75)

    plt.legend(loc='upper center', ncol=5, frameon=True, bbox_to_anchor=(0.5, 1.1), fancybox=True, shadow=True, fontsize=10)
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    plt.xlabel('Condition Class', size=14)
    plt.xticks(rotation=45, ha = 'right', size=20)
    plt.ylabel('Number of Landscapes', size=14)
    plt.yticks(size=20)
    plt.title(nameDict.get(var), size = 18, y=1.1)
    plt.tight_layout()

    print("saving condition/threat class figure for {}...".format(var))

    outName = var + "_barchart_byClass"
    outFile = os.path.join(OUTPUT_PATH, outName)
    plt.savefig(outFile, dpi=300)

    plt.close()
