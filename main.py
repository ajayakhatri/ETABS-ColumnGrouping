# %% [markdown]
# This code does following:
# 1. Connects to a running ETABS window
# 2. Reads reinforcement data from data.xlsx.
# 3. Selects columns in ground floor.
# 4. Finds columns lying just above it on upper floors.
# 5. Assign Groups to the columns based on percentage rebar of column in ground floor.

# %%
import comtypes.client
import pandas as pd

# Connects To Etabs
ETABSObject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
SapModel = ETABSObject.SapModel


# %%
df_data=pd.read_excel("data.xlsx")

# %%
#getColumns returns list of frames with Design Orientation 1
def getColumns(frames): 
        columns=[] 
        for frame in frames:
                if SapModel.FrameObj.GetDesignOrientation(frame)[0]==1:
                        columns.append(frame)     
        return columns

def startPoint(x):
    return SapModel.FrameObj.GetPoints(x)[0]

def location(x):
    return SapModel.PointObj.GetCoordCartesian(x)

# %%
allFrames =SapModel.FrameObj.GetAllFrames()[1]
fristStoryFrames = SapModel.FrameObj.GetNameListOnStory("Story1")[1]

allColumns=getColumns(allFrames)
fristStoryColumns=getColumns(fristStoryFrames)

# %%
# sameLocationColumn returns a list containing list of columns lying in same (x,y) coordinate
def sameLocationColumn(columns, groundColumns):
    groups=[]
    for groundColumn in groundColumns:
        group=[groundColumn]
        for column in columns:
            if not column in groundColumns:
                if location(startPoint(column))[0]==location(startPoint(groundColumn))[0] and location(startPoint(column))[1]==location(startPoint(groundColumn))[1]:
                     group.append(column)
        groups.append(group)
    return groups

# %%
sameLocationColumn=sameLocationColumn(allColumns,fristStoryColumns)

# %%
#adds rebarPercent Column in the data table
def rebarPercent(df_data):
    percentRebar=[]
    if not "rebarPercent" in df_data:
        for i in df_data["As"]:
            percentRebar.append(i/(500*500)*100)     #length Of Square Beam= 500mm
        df_data['rebarPercent']= percentRebar
    else:
        print("ERROR COLUMN ALREADY EXIST!!!")

# %%
rebarPercent(df_data)

# %%
# assigns column in groups based on provided ranges of percentage reinforcement
def rebarGroup(df_data, lower, upper,sameLocationColumn, groupName):
    group=[]
    index=0
    count=0
    SapModel.GroupDef.SetGroup(groupName)
    for a in df_data["UniqueName"]:
        i=str(a)
        if lower < df_data["rebarPercent"][index] and df_data["rebarPercent"][index]<= upper:
            for x in sameLocationColumn:
                if i == x[0]:
                    for j in x:
                        SapModel.FrameObj.SetGroupAssign(str(j), groupName)
                        SapModel.FrameObj.SetSelected(str(j),True)
                        group.append(str(j))
        index=index+1
    print("DONE")
    return list(dict.fromkeys(group))

# %%

low=0
high=0.824353912301962
groupname="C501"
C501=rebarGroup(df_data,low,high,sameLocationColumn,groupname)

# %%
low=0.824353912301962
high=1.46775208775715
groupname="C502"
C502=rebarGroup(df_data,low,high,sameLocationColumn,groupname)


# %%
low=1.46775208775715
high=1.75049542658023
groupname="C503"
C503=rebarGroup(df_data,low,high,sameLocationColumn,groupname)


# %%

low=1.75049542658023
high=2.01061929829747
groupname="C504"
C504=rebarGroup(df_data,low,high,sameLocationColumn,groupname)


