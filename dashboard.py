import streamlit as st 
import matplotlib.pyplot as plt
import plotly.express as px
import pandas as pd
from io import BytesIO
import numpy as np
import time
import os
import datetime as dt
import textwrap


st.set_page_config(layout="wide", page_title="FIELDSA DASHBOARD")
with open("header.html", "r") as head:
    st.markdown(head.read(), unsafe_allow_html=True)

    
if "menu_sidebar" not in st.session_state:
    st.session_state.menu_sidebar = "WorkOrder Chart"
st.sidebar.title("Sidebar Menu")
if st.sidebar.button("WorkOrder Chart"):
    st.session_state.menu_sidebar = "WorkOrder Chart"
if st.sidebar.button("Status Chart"):
    st.session_state.menu_sidebar = "Status Chart"
if st.sidebar.button("SLA Summary"):
    st.session_state.menu_sidebar = "SLA Summary"
menu_sidebar= st.session_state.menu_sidebar

 
uploaded=st.file_uploader("Upload Excel File", type=["xlsx", "csv"]) 

if uploaded is not None:
    up2webtime= pd.Timestamp.now()
    @st.cache_data
    def load(file):
        if file.name.endswith(".csv"):
            return {"WorkOrder": pd.read_csv(file)}
        else:
            return pd.read_excel(file, sheet_name=None)
    exceldata=load(uploaded)
    df= exceldata.get("WorkOrder")
    if df is not None:
        df["uptime"]= up2webtime
    
    
    

    
    def parse(durasi):
        if durasi is None or durasi in ("N/A", "") or pd.isna(durasi):
            return None
        bagian=str(durasi).split(":")
        if len(bagian) != 4:
            return None
        days, hours, minutes, seconds= map(int, bagian)
        return days*24*60+ hours*60+minutes

    def klasifikasi(durasi, tipe):
        menit=parse(durasi)
        if menit is None:
            return "N/A"
        if tipe== "short":
            if menit <= 15:
                return "<15 Menit"
            elif menit <=30:
                return "15-30 Menit"
            else:
                return ">30 Menit"
        if tipe== "long":
            jam= menit/60
            if jam <=8:
                return "<8 Jam"
            elif jam <=16:
                return "8-16 Jam"
            elif jam <=24:
                return "16-24 Jam"
            else:
                return ">24 Jam"
        return "N/A"            

    def durasi(awal,akhir,ts):  
        if awal in ts and akhir in ts and pd.notna(ts[awal]) and pd.notna(ts[akhir]):
            count= ts[akhir]-ts[awal]
            totalsec= int(pd.Timedelta(count).total_seconds())
            days, remains= divmod(totalsec, 86400)
            hours, remains= divmod(remains, 3600)
            minutes, seconds= divmod(remains, 60)
            return f"{days:02}:{hours:02}:{minutes:02}:{seconds:02}"
        else:
            return ""   
 

    def exportfile(uploadedfile):
        if not uploadedfile.name.lower().endswith(".xlsx"):
            return None, None
        
        xcel= pd.ExcelFile(uploadedfile)
        pake= {"HistoryWorkOrder", "WorkOrder", "Rca"}
        if not pake.issubset(set(xcel.sheet_names)):
            return None, None
        
        dfcalc= pd.read_excel(xcel, sheet_name="HistoryWorkOrder")
        dfcalc.columns= dfcalc.columns.str.strip()
        dfcalc = dfcalc.rename(columns={
            'WorkOrderNumber':'WONumber',
            'WorkOrderStatusItem' : 'status',
            'Modified' : 'timestamp'  
        })

        dfcalc['status'] = dfcalc['status'].astype(str).str.strip()
        dfcalc['timestamp'] = pd.to_datetime(dfcalc['timestamp'])
        dfcalc= dfcalc.dropna(subset=['WONumber', 'timestamp'])
        dfcalc= dfcalc.sort_values(['WONumber', 'timestamp'])

        urutanstatus=[
            'Open',
            'Assign to dispatch external',
            'Assign to technician',
            'Accept',
            'Travel',
            'Arrive',
            'On Progress',
            'Done',
            'Complete with note request',
            'Postpone Request',
            'Complete',
            'Complete with note approve',
            'Postpone'
        ]

        applyurutanstatus= set(urutanstatus)

        hasil= []
        statusdur=[]

        for wo, group in dfcalc.groupby('WONumber'):
            statusnump= group['status'].to_numpy()
            timenump= group['timestamp'].to_numpy()
            baris= {'WONumber': wo}
            timestemp={}
            lasttime=pd.Timestamp.min
            baris['Anomali']= 'Abnormal' if pd.Series(statusnump).duplicated().any() else 'Normal'

            for stat in urutanstatus:
                filtertimenstat= (statusnump==stat)&(timenump>lasttime)
                if filtertimenstat.any():
                    if stat== 'Open':
                        selectedindex= np.argmax(filtertimenstat)
                    else:
                        selectedindex= np.where(filtertimenstat)[0][-1]
                    timestemp[stat]=timenump[selectedindex]
                    lasttime= timenump[selectedindex]
                else:
                    timestemp[stat]= pd.NaT
            for stat in set(statusnump)-applyurutanstatus:
                timestemp[stat]= timenump[statusnump== stat].max()
            
            baris.update(timestemp)
            hasil.append(baris)
            
            
            d1 = durasi('Open', 'Assign to dispatch external', timestemp)
            d2 = durasi('Assign to dispatch external', 'Assign to technician', timestemp)
            d3 = durasi('Assign to technician', 'Accept', timestemp)
            d4 = durasi('Accept', 'Done', timestemp)
            d5 = durasi('Done', 'Complete', timestemp)
            
            statusdur.append({
                'WONumber': wo, 
                'Open - Assign to dispatch external': d1,
                'SLA Open-Dispatch External': klasifikasi(d1, "short"),
                'Assign to dispatch external - Assign to technician': d2,
                'SLA Dispatch External -Technician': klasifikasi(d2, "short"),
                'Assign to technician - Accept': d3,
                'SLA Technician-Accept': klasifikasi(d3, "short"),
                'Accept - Done': d4,
                'SLA Accept-Done': klasifikasi(d4, "long"),
                'Done - Complete': d5,
                'SLA Done-Complete': klasifikasi(d5, "short")
            })

        final=pd.DataFrame(hasil).merge(pd.DataFrame(statusdur), on='WONumber', how='left' )

        addcols= [ 'WorkOrderNumber', 'ReferenceCode', 'WorkOrderTypeName', 'DivisionName', 'WorkOrderStatusItem', 'Reason',
            'CustomerId', 'CustomerName', 'Cid', 'CircuitId', 'EndCustomerName', 'SubRegion',
            'City', 'DeviceAllocation', 'VendorName', 'DispatcherName', 'TechnicianName']

        df2= pd.read_excel(xcel, sheet_name='WorkOrder', usecols=addcols).rename(columns={'WorkOrderNumber':'WONumber'}).reset_index(drop=True)
        df2['SubRegion'] =df2['SubRegion'].astype(str).str.strip().str.title()
        df2['WorkOrderStatusItem']= df2['WorkOrderStatusItem'].astype(str).str.strip().str.title()

        regionmap = {
            'Central Java': 'Central',
            'Jabodetabek': 'Central',
            'West Java': 'Central',
            'Bali': 'East',
            'East Java': 'East',
            'Kalimantan': 'East',
            'Sulawesi': 'East',
            'Internasional': 'Internasional',
            'Kepulauan Riau': 'West',
            'Northern Sumatera': 'West',
            'Southern Sumatera': 'West'
        }

        statusreportmap= {
            'Open':'OPEN',
            'Assign To Dispatch External':'ONPROGRESS',
            'Complete With Note Approve':'COMPLETE',
            'Assign To Technician':'ONPROGRESS',
            'Complete':'COMPLETE',
            'Accept':'ONPROGRESS',
            'Travel' :'ONPROGRESS',
            'Arrive':'ONPROGRESS',
            'On Progress':'ONPROGRESS',
            'Return':'ONPROGRESS',
            'Done':'COMPLETE',
            'Work Order Confirmation Approve':'COMPLETE',
            'Complete With Note Request':'COMPLETE',
            'Complete With Note Reject':'ONPROGRESS',
            'Postpone Request':'POSTPONE',
            'Postpone':'POSTPONE',
            'Sms Integration Failed':'INTEGRATION FAILED',
            'Revise':'ONPROGRESS',
            'Return By Technician':'ONPROGRESS',
            'Postpone Is Revised':'POSTPONE',
            'Return Is revised':'ONPROGRESS',
            'Provisioning In Progress':'ONPROGRESS',
            'Provisioning Success':'ONPROGRESS',
            'Posted To Ax Integration Failed':'INTEGRATION FAILED',
            'Provisioning Failed':'INTEGRATION FAILED',
            'Cancel Work Order':'CANCEL',
            'Posted To Ax Integration Success':'COMPLETE'
        }
        final= final.merge(df2, on='WONumber', how='left')
        final['Region']= final['SubRegion'].map(regionmap).fillna('N/A')
        final['StatusReport']= final['WorkOrderStatusItem'].map(statusreportmap).fillna('N/A')
        subregionindex= final.columns.get_loc('SubRegion')
        final.insert(subregionindex, 'Region', final.pop('Region'))
        wostatusindex= final.columns.get_loc('WorkOrderStatusItem')
        final.insert(wostatusindex, 'StatusReport', final.pop('StatusReport'))

        df3= pd.read_excel(xcel, sheet_name='Rca', usecols=['WorkOrderNumber', 'UpTime']).rename(columns={'WorkOrderNumber':'WONumber'}).reset_index(drop=True)
        final= final.merge(df3, on='WONumber', how='left')
        final['UpTime'] = final['UpTime'].fillna('N/A')



        urutanstatuswo= [
            'WONumber','Anomali','Open','Assign to dispatch external','Assign to technician','Accept',
            'Travel','Arrive','On Progress','Done','Work Order Confirmation Approve','Complete',
            'Complete with note approve','Complete with note request','Complete with note reject',
            'Postpone Request','Postpone is Revised','Postpone','SMS Integration Failed','Return',
            'Return by Technician','Revise','Return is revised','Provisioning In Progress','Provisioning Success',
            'Posted to AX Integration Failed','Posted to AX Integration Success','Provisioning Failed','Cancel Work Order',
            'ReferenceCode','WorkOrderTypeName','DivisionName','StatusReport','WorkOrderStatusItem','Reason','CustomerId',
            'CustomerName','Cid','CircuitId','EndCustomerName','Region','SubRegion','City','DeviceAllocation',
            'VendorName','DispatcherName','TechnicianName','UpTime',
            'Open - Assign to dispatch external','SLA Open-Dispatch External',
            'Assign to dispatch external - Assign to technician','SLA Dispatch External -Technician',
            'Assign to technician - Accept','SLA Technician-Accept','Accept - Done','SLA Accept-Done',
            'Done - Complete','SLA Done-Complete'
        ]

        final= final[[col for col in urutanstatuswo if col in final.columns]]

        buffer= BytesIO()
        with pd.ExcelWriter(buffer, engine= "xlsxwriter") as writer:
            final.to_excel(writer, index= False, sheet_name= "SLA")
        buffer.seek(0)
        suggestname = f"AllTaskList_FIELDSA_{pd.Timestamp.now():%Y%m%d_%H%M%S}.xlsx"
        return buffer, suggestname
    
        # st.write("## Status Duration & SLA Calculation")
        # st.caption("Click the button below to calculate the status duration, status SLA and export as Excel")  
        # exportbutton, fillerexpbutton1, fillerexpbutton2= st.columns([1, 2, 3])
        # with exportbutton:
        #     if st.button("Export to Excel", type="primary", use_container_width= True):
        #         with st.spinner("Processing..."):
        #             out, filename= exportfile(uploaded)
        #         if out is not None:
        #             st.success("Export ready")
        #             st.download_button("Download Excel", data= out.getvalue(), file_name= filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True )
        #         else:
        #             st.warning("Export Failed")
                                






    
    for col in df.columns:
        if df[col].dtype== object:
            df[col]= df[col].astype(str).str.strip().str.title().replace("Nan", "N/A")
    
    
    df["Created"]= pd.to_datetime(df["Created"], errors="coerce")
    df= df.dropna(subset=["Created"])
    firstdate= df["Created"].min().date()
    lastdate= df["Created"].max().date()
    rangetgl= st.date_input(
        "Select data range",
        (firstdate,lastdate),
        min_value= firstdate,
        max_value= lastdate
    )
    if len(rangetgl)== 2:
        tglawal,tglakhir= rangetgl
        df= df[(df["Created"]>=pd.to_datetime(tglawal)) &
               (df["Created"]<=pd.to_datetime(tglakhir)+pd.Timedelta(days=1)-pd.Timedelta(seconds=1))] 
    regionmap={
        "Bali": "East",
        "Central Java": "Central",
        "East Java": "East",
        "Jabodetabek":"Central",
        "West Java": "Central",
        "Kalimantan":"East",
        "Sulawesi":"East",
        "Internasional": "International",
        "Kepulauan Riau": "West",
        "Northern Sumatera":"West",
        "Southern Sumatera":"West"
    }

    if "SubRegion" in df.columns:
        df["Region"]= df["SubRegion"].map(regionmap).fillna("Unknown")
    if df is not None:
        col1,col2, col3=st.columns([1,2,2])
        col1.metric("Total WO Number", len(df))
        try:
            identify= pd.to_datetime(uploaded.name.split("_")[1].split(".")[0], format="%Y%m%d%H%M%S")
            lastupdate= identify.strftime("%d %B %Y %H:%M:%S")
        except:
            lastupdate= "Date not valid"
    
        col2.metric("Date pulled from FIELDSA", lastupdate)
        last_input_refresh= df["uptime"].max().strftime("%d %B %Y %H:%M:%S")
        col3.metric("Last Input/Refresh", last_input_refresh)
        
        dfstat= df.copy() 
        
        
        
        
        
        
        
        
        if menu_sidebar == "WorkOrder Chart":
            st.divider()
            st.write("## WorkOrder Chart")
            location= ["Region", "SubRegion", "City [All]", "City [Top 10]"]
            ubahheader={
                "WorkOrderTypeName":"Work Order Type",
                "DivisionName": "Division",
                "Region": "Region",
                "SubRegion":"Sub Region",
                "City": "City"
            }
            
            filter1,filter2, filter3=st.columns(3)
            
            if "locfilter" not in st.session_state:
                st.session_state.locfilter= "Region"
            if "divfilter" not in st.session_state:
                st.session_state.divfilter= ["Broadband", "Lms", "Fiberisasi"]
            if "tipefilter" not in st.session_state:
                st.session_state.tipefilter= ["Troubleshoot", "Activation"]
            
            
            with filter1:
                piliharea = st.selectbox("Location",options=["Region", "SubRegion", "City [All]", "City [Top 10]"],format_func= lambda x: "Sub Region" if x== "SubRegion" else x, key= "location1")
                loccol = "City" if "City" in piliharea else piliharea 
                if loccol in df.columns:
                    if piliharea == "City [Top 10]":
                        yangdisaji = df["City"].value_counts().nlargest(10).index.tolist()
                    else:
                        yangdisaji = sorted(df[loccol].dropna().astype(str).unique().tolist())
                else:
                    yangdisaji = []

                filter_dalamarea = st.multiselect(f"Select {loccol}", options=yangdisaji, default=yangdisaji, key=f"multisel1_{loccol}")
                if filter_dalamarea:
                    df=df[df[loccol].astype(str).isin(filter_dalamarea)]
                else:
                    st.warning("gaada filter yang dipilih")
                
                        # loc= st.selectbox("Location", options=location, format_func= lambda x: "Sub Region" if x== "SubRegion" else x, key= "locfilter")
                        # loccol= "City" if "City" in loc else loc
                        # if loc=="City [Top 10]":
                        #     top10= df["City"].value_counts().nlargest(10).index.tolist()
                        #     df= df[df["City"].isin(top10)]
            
            with filter2:
                st.markdown("WO Type")
                with st.expander("WO Type"):
                    tipewo= ['Troubleshoot', 'Activation']
                    jumlahtipe= df["WorkOrderTypeName"].value_counts().to_dict()
                    
                    tipeygdipilih=[]
                    
                    for r in tipewo:
                        itungg= jumlahtipe.get(r,0)
                        if st.checkbox(f"{r} ({itungg})", value= r in st.session_state.tipefilter, key=f"{r}" ):
                            tipeygdipilih.append(r) 
                    st.session_state.tipefilter= tipeygdipilih              
                    
            with filter3:
                st.markdown("Division")
                with st.expander("Division"):
                    divwo= ["Broadband","Lms","Fiberisasi"]
                    jumlahperdiv= df["DivisionName"].value_counts().to_dict()
                    divisiterpilih=[]
                    for c in divwo:
                        itung= jumlahperdiv.get(c,0)
                        if st.checkbox(f"{c} ({itung})", value= c in st.session_state.divfilter, key=f"{c}"):
                            divisiterpilih.append(c)
                    st.session_state.divfilter= divisiterpilih
    
            if divisiterpilih:
                df= df[df["DivisionName"].isin(divisiterpilih)]
            else:
                st.warning("No division type chosen")
                
            if tipeygdipilih:
                df= df[df["WorkOrderTypeName"].isin(tipeygdipilih)]
            else:
                st.warning("No Workorder type chosen")
    
            
            
            if loccol in df.columns:
                normcol = df[loccol].astype(str)
                itungisi = normcol.value_counts().reset_index()
                itungisi.columns = [loccol, "Count"]

                bar = px.bar(
                    itungisi,
                    x= loccol,
                    y= "Count",
                    color= loccol,
                    title= f"Based on {ubahheader.get(loccol, loccol)}",
                    labels= {loccol: ubahheader.get(loccol,loccol), "Count": "Jumlah"}
                )
                st.plotly_chart(bar, use_container_width=True)
            else:
                st.warning("gaada kolom yang bisa dipakai")
            tambahcols = [loccol, "CustomerName", "VendorName", "Reason"]
            colkotak = st.columns(len(tambahcols))
            for index, namadicol in enumerate(tambahcols):
                with colkotak[index]:
                    if namadicol in df.columns:
                        addcoldata = df[namadicol].astype(str).value_counts().reset_index()
                        addcoldata.columns = [namadicol, "Amount"]
                        total= pd.DataFrame({namadicol: ["Total"], "Amount": [addcoldata["Amount"].sum()]})
                        addcoldata= pd.concat([addcoldata, total])
                        addcoldata.index = range(1, len(addcoldata) + 1)
                        st.markdown(f"**{namadicol}**")
                        st.dataframe(addcoldata)
                    else:
                        st.warning(f"Kolom {namadicol} tidak ditemukan")           
                    
        
                
                
                
                
                
                
                
        if menu_sidebar == "Status Chart":
            st.divider()   
            st.write("## Work Order Status Chart")   
            
            dfstat["WorkOrderStatusItem"] = dfstat["WorkOrderStatusItem"].astype(str).str.strip().str.title()
            location= ["Region", "SubRegion", "City [All]", "City [Top 10]"]
            ubahheader={
                "WorkOrderTypeName":"Work Order Type",
                "DivisionName": "Division",
                "Region": "Region",
                "SubRegion":"Sub Region",
                "City": "City"
            }
            statusmap={
                "Open": "OPEN",
                "Assign To Dispatch External": "ONPROGRESS",
                "Complete With Note Approve": "COMPLETE",
                "Assign To Technician": "ONPROGRESS",
                "Complete": "COMPLETE",
                "Accept": "ONPROGRESS",
                "Travel": "ONPROGRESS",
                "Arrive": "ONPROGRESS",
                "On Progress": "ONPROGRESS",
                "Return": "ONPROGRESS",
                "Done": "COMPLETE",
                "Work Order Confirmation Approve": "COMPLETE",
                "Complete With Note Request": "COMPLETE",
                "Complete With Note Reject": "ONPROGRESS",
                "Postpone Request": "POSTPONE",
                "Postpone": "POSTPONE",
                "Sms Integration Failed": "INTEGRATION FAILED",
                "Revise": "ONPROGRESS",
                "Return By Technician": "ONPROGRESS",
                "Postpone Is Revised": "POSTPONE",
                "Return Is Revised": "ONPROGRESS",
                "Provisioning In Progress": "ONPROGRESS",
                "Provisioning Success": "ONPROGRESS",
                "Posted To Ax Integration Failed": "INTEGRATION FAILED",
                "Provisioning Failed": "INTEGRATION FAILED",
                "Cancel Work Order": "CANCEL",
                "Posted To Ax Integration Success": "COMPLETE" }
            
            dfstat["StatusReport"]= dfstat["WorkOrderStatusItem"].map(statusmap).fillna("N/A")
            
            locfilter2,statfilter=st.columns(2)
            
            if "location2" not in st.session_state:
                st.session_state.location2= "Region"
            if "status" not in st.session_state:
                st.session_state.status= ["OPEN", "COMPLETE", "ONPROGRESS", "POSTPONE", "INTEGRATION FAILED", "CANCEL"]
    
            with locfilter2:
                loc= st.selectbox("Location", options=location, format_func= lambda x: "Sub Region" if x== "SubRegion" else x, key= "location2")
                loccol= "City" if "City" in loc else loc
                if loccol in dfstat.columns:
                    if loc=="City [Top 10]":
                        pilihh= dfstat["City"].value_counts().nlargest(10).index.tolist()
                    else:
                        pilihh= sorted(dfstat[loccol].dropna().astype(str).unique().tolist())
                else:
                    pilihh=[]
                filter_dalamarea= st.multiselect(f"Select {loccol}", options= pilihh, default=pilihh, key=f"multisel2_{loccol}")
                if filter_dalamarea:
                    dfstat=dfstat[dfstat[loccol].astype(str).isin(filter_dalamarea)]
                else:
                    st.warning("gaada filter yang dipilih")

            with statfilter:
                st.markdown("Status")
                with st.expander("Status"):
                    statwo= ["OPEN", "COMPLETE", "ONPROGRESS", "POSTPONE", "INTEGRATION FAILED","CANCEL"]
                    jumlahperstat= dfstat["StatusReport"].value_counts().to_dict()
                    
                    statusterpilih=[]
                
                    for s in statwo:
                        itung= jumlahperstat.get(s,0)
                        if st.checkbox(f"{s} ({itung})", value= (s in st.session_state.status), key=f"status_{s}"):
                            statusterpilih.append(s)
                    st.session_state.status= statusterpilih
                    
            if statusterpilih:
                filtereddf= dfstat[dfstat["StatusReport"].isin(statusterpilih)]
            else:
                st.warning("No status picked")
                filtereddf = pd.DataFrame()
            
            if not filtereddf.empty:
                statusgroup=(filtereddf.groupby([loccol, "StatusReport"]).size().reset_index(name="Count"))
                statusbar= px.bar(
                        statusgroup,
                        x= loccol,
                        y= "Count",
                        color= "StatusReport",
                        title= f"Based on {ubahheader.get(loccol, loccol)}",
                )
                
                st.plotly_chart(statusbar, use_container_width=True)
            
                statussummary= statusgroup.pivot_table(index=loccol, columns= "StatusReport", values= "Count", fill_value=0).reset_index()
                statussummary["Total"]= statussummary.loc[:, statussummary.columns != loccol].sum(axis=1)
                totalperstatus= statussummary.loc[:, statussummary.columns != loccol].sum(axis=0)
                totalbaris=pd.DataFrame([totalperstatus])
                
                totalbaris[loccol]= "Total"
                statussummary= pd.concat([statussummary, totalbaris], ignore_index= True)
                statussummary.index= range(1, len(statussummary)+1)


                st.dataframe(statussummary)
            else:
                st.warning("No data available")






        if menu_sidebar == "SLA Summary":
            st.divider()
            st.write("## Status Report SLA WorkOrder Summary")
            statusmap1={
            "Open": "OPEN",
            
            "Assign To Technician": "ONPROGRESS",
            "Accept": "ONPROGRESS",
            "Travel": "ONPROGRESS",
            "Arrive": "ONPROGRESS",
            "On Progress": "ONPROGRESS",
            "Return": "ONPROGRESS",
            "Assign To Dispatch External": "ONPROGRESS",
            "Complete With Note Reject": "ONPROGRESS",
            "Revise": "ONPROGRESS",
            "Return By Technician": "ONPROGRESS",
            "Postpone Is Revised": "POSTPONE",
            "Return Is Revised": "ONPROGRESS",
            "Provisioning In Progress": "ONPROGRESS",
            "Provisioning Success": "ONPROGRESS",
            
            "Complete With Note Approve": "COMPLETE",
            "Complete": "COMPLETE",
            "Done": "COMPLETE",
            "Work Order Confirmation Approve": "COMPLETE",
            "Complete With Note Request": "COMPLETE",
            "Posted To Ax Integration Success": "COMPLETE",
            
            "Postpone Request": "POSTPONE",
            "Postpone": "POSTPONE",
            
            "Sms Integration Failed": "INTEGRATION FAILED",
            "Posted To Ax Integration Failed": "INTEGRATION FAILED",
            "Provisioning Failed": "INTEGRATION FAILED",
            
            "Cancel Work Order": "CANCEL",}

            if isinstance(exceldata, dict) and "WorkOrder" in exceldata:
                sladf= exceldata["WorkOrder"].copy()
            elif"df" in locals():
                sladf= df.copy()
            else:
                sladf= pd.DataFrame()
            
            for s in sladf.columns:
                if sladf[s].dtype== object:
                    sladf[s]= sladf[s].astype(str).str.strip()
            if "Region" not in sladf.columns:
                sladf["Region"]= sladf["SubRegion"].map(regionmap).fillna(sladf["SubRegion"])
        
            location3= ["Region", "SubRegion", "City [All]", "City [Top 10]"]
            kolsla1, kolsla2, kolsla3 = st.columns(3)
            st.session_state.location3= "Region"
            st.session_state.divfilter2= ["Broadband", "Lms", "Fiberisasi"]
            with kolsla1:
                area= st.selectbox("Location", options=location3, format_func= lambda x: "Sub Region" if x== "SubRegion" else x, key= "locfilter3")
                loccol= "City" if "City" in area else area
                if loccol in sladf.columns:
                    if area== "City [Top 10]":
                        tampilkan= sladf["City"].value_counts().nlargest(10).index.tolist()
                    else:
                        tampilkan= sorted(sladf[loccol].dropna().astype(str).unique().tolist())
                else:
                    tampilkan=[]
                filter_diarea= st.multiselect(f"Select {loccol}", options=tampilkan, default=tampilkan, key=f"multisel3_{loccol}")
                if filter_diarea:
                    sladf=sladf[sladf[loccol].astype(str).isin(filter_diarea)]
                else:
                    st.warning("gaada filter yang dipilih")
            with kolsla2:
                st.markdown("WO Type")
                with st.expander("WO Type"):
                    pilihantipe= sorted(sladf["WorkOrderTypeName"].dropna().unique().tolist())if "WorkOrderTypeName" in sladf.columns else []
                    tipeyangdipilih=[]
                    for p in pilihantipe:
                        defaultset= (p in st.session_state.tipefilter) if "tipefilter" in st.session_state else False
                        countlabel= int(sladf["WorkOrderTypeName"].value_counts().get(p,0)) if "WorkOrderTypeName" in sladf.columns else 0
                        if st.checkbox(f"{p}({countlabel})", value=defaultset, key=f"wotypesla_{p}"):
                            tipeyangdipilih.append(p)
                    st.session_state.tipefilter=tipeyangdipilih
                
            with kolsla3:
                st.markdown("Division")
                with st.expander("Division"):
                    divwo= ["Broadband","Lms","Fiberisasi"]
                    jumlahperdiv= sladf["DivisionName"].value_counts().to_dict()
                    divisiterpilih=[]
                    for c in divwo:
                        itung= jumlahperdiv.get(c,0)
                        if st.checkbox(f"{c} ({itung})", value= c in st.session_state.divfilter2, key=f"{c}_sla"):
                            divisiterpilih.append(c)
                    st.session_state.divfilter2= divisiterpilih
                    if st.session_state.divfilter2:
                        sladf=sladf[sladf["DivisionName"].astype(str).isin(st.session_state.divfilter2)].copy()
                    else:
                        st.warning("gaada filter yang dipilih")
            if tipeyangdipilih:
                if "WorkOrderTypeName" in sladf.columns:
                    sladf= sladf[sladf["WorkOrderTypeName"].isin(tipeyangdipilih)].copy()
                else:
                    st.warning("gaada wotypename di sladf")
                    sladf=pd.DataFrame()
                if area=="City [Top 10]":
                    if "City" in sladf.columns:
                        top10= sladf["City"].value_counts().nlargest(10).index.tolist()
                        sladf= sladf[sladf["City"].isin(top10)].copy()
                    else:
                        st.warning("kolom city gaditemuin") 
                        
                        
                        
                if (isinstance(exceldata, dict)and "HistoryWorkOrder" in exceldata):
                    stathistory= exceldata["HistoryWorkOrder"].copy()
                    stathistory.columns=stathistory.columns.astype(str).str.strip()
                    his_wo= "WorkOrderNumber" if "WorkOrderNumber" in stathistory.columns else None
                    his_stat= "WorkOrderStatusItem" if "WorkOrderStatusItem" in stathistory.columns else None
                    his_time= "Modified" if "Modified" in stathistory.columns else None
                    if all([his_wo, his_stat, his_time]):
                        scannedwo=(sladf[["WorkOrderNumber","WorkOrderStatusItem"]].dropna().copy())
                        scannedwo["WorkOrderNumber"]= scannedwo["WorkOrderNumber"].astype(str).str.strip()
                        scannedwo["wonum_key"]= scannedwo["WorkOrderNumber"].str.title()
                        scannedwo["statusnormalized"]= scannedwo["WorkOrderStatusItem"].astype(str).fillna("").str.strip().str.title()
                        simpen_wonum_key= set(scannedwo["wonum_key"].unique())
                        stathistory["wonum_key"]=(stathistory[his_wo].astype(str).fillna("").str.strip().str.title())
                        his_subset= stathistory[stathistory["wonum_key"].isin(simpen_wonum_key)].copy()
                        his_subset["WorkOrderNumber"]= his_subset[his_wo].astype(str).str.strip()
                        his_subset["StatusWO"]= his_subset[his_stat].astype(str).str.strip()
                        his_subset["StatusTimestamp"]= pd.to_datetime(his_subset[his_time],errors="coerce")
                        his_subset= his_subset.dropna(subset=["WorkOrderNumber","StatusWO","StatusTimestamp"])
                        if his_subset.empty:
                            st.write("history stlh normalisasi kosong") 
                        else:
                            openonly= (his_subset[his_subset["StatusWO"].str.title() == "Open"].groupby("WorkOrderNumber", as_index= False).agg({"StatusTimestamp":"min"}).rename(columns={"StatusTimestamp":"open_c"}))
                            others= (his_subset[his_subset["StatusWO"].str.title() != "Open"].groupby(["WorkOrderNumber","StatusWO"], as_index= False).agg({"StatusTimestamp":"max"}).rename(columns={"StatusTimestamp":"status_c"}))
                            gabung= others.merge(openonly, on="WorkOrderNumber", how="left")
                            showopen=openonly.copy()
                            showopen["StatusWO"]= "Open"
                            showopen["status_c"]= pd.NaT
                            showopen["terpilih_c"]= showopen["open_c"]
                            gabung["terpilih_c"]= gabung["status_c"]
                            kol=["WorkOrderNumber","StatusWO","status_c","open_c","terpilih_c"]
                            gabungsemua=pd.concat([showopen.loc[:, kol],gabung.loc[:, kol]], ignore_index=True, sort=False)
                            def statusreportmap(a):
                                if a in statusmap1:
                                    return statusmap1[a]
                                return statusmap1.get(str(a).title(), "OTHER")
                            gabungsemua["StatusReport"]=  gabungsemua["StatusWO"].apply(statusreportmap)
                            gabungsemua= gabungsemua[gabungsemua["StatusReport"].isin(["OPEN","ONPROGRESS","POSTPONE","COMPLETE", "INTEGRATION FAILED","CANCEL"])].copy()
                            gabungsemua["wonum_key"]= gabungsemua["WorkOrderNumber"].astype(str).str.strip().str.title()                     
                            gabungsemua["statusnormalized"]= gabungsemua["StatusWO"].astype(str).fillna("").str.strip().str.title()
                            tergabung= gabungsemua.merge(scannedwo[["wonum_key", "statusnormalized"]], on="wonum_key", how="inner", suffixes=("", "wo"))
                            tergabung=tergabung[tergabung["statusnormalized"]==tergabung["statusnormalizedwo"]].copy()

                            if tergabung.empty:
                                st.write("gaada yg match antara history stat dan stat excel")       
                            else:
                                tergabung["WorkOrderNumber"]= tergabung["WorkOrderNumber"].astype(str).str.strip().str.upper()
                                sladf["WorkOrderNumber"]= sladf["WorkOrderNumber"].astype(str).str.strip().str.upper()
                                tergabung= tergabung.merge(sladf[["WorkOrderNumber","uptime"]], on="WorkOrderNumber", how="left")
                                tergabung["duration"]=((tergabung["uptime"]-tergabung["terpilih_c"]).dt.total_seconds()/3600)
                                completeonly= tergabung["StatusReport"]=="COMPLETE"
                                tergabung.loc[completeonly, "duration"]= ((tergabung.loc[completeonly, "status_c"]-tergabung.loc[completeonly, "open_c"]).dt.total_seconds()/3600)
                                def slaoptions(hour):
                                    if hour<=6:
                                        return "0-6 Jam"
                                    elif hour<=12:
                                        return "6-12 Jam"
                                    elif hour<=18:
                                        return "12-18 Jam"
                                    elif hour<=24:
                                        return "18-24 Jam"
                                    elif pd.isna(hour):
                                        return None
                                    else:
                                        return ">24 Jam"
                            
                                kol_area=  area if area != "City [All]" and area != "City [Top 10]" else "City"
                                if kol_area in sladf.columns:
                                    tergabung= tergabung.merge(sladf[["WorkOrderNumber", kol_area]].drop_duplicates(), on="WorkOrderNumber", how="left")
                                else:
                                    st.write("ganemu kol_area")
                                tergabung["slaoptions"]=tergabung["duration"].apply(slaoptions)
                                tergabung_valid= tergabung.dropna(subset=["slaoptions"]).copy()
                                slagroup=( tergabung_valid.groupby([kol_area, "StatusReport", "slaoptions"]).agg({"WorkOrderNumber": "nunique"}).reset_index())
                                                                                                                
                                # st.dataframe(slagroup)
                                # st.write(tergabung)
                                if "WorkOrderNumber" in slagroup.columns:
                                    slagroup=slagroup.rename(columns={"WorkOrderNumber":"Count"})
                                slagroup["Count"]= slagroup["Count"].astype(int)
                                urutanstatus=["OPEN","ONPROGRESS","POSTPONE","COMPLETE", "INTEGRATION FAILED", "CANCEL"]
                                urutansla=["0-6 Jam", "6-12 Jam", "12-18 Jam", "18-24 Jam", ">24 Jam"]
                                pivot= slagroup.pivot_table(index=[kol_area,"slaoptions"],columns="StatusReport",values="Count",aggfunc="sum", fill_value=0)
                                for u in urutanstatus:
                                    if u not in pivot.columns:
                                        pivot[u]= 0
                                pivot= pivot.reindex(columns=urutanstatus, fill_value=0)
                                pivot["TOTAL"]= pivot.sum(axis=1)
                                pivot.index.set_names([kol_area, "SLA"], inplace= True)
                                areasum= pivot.groupby(level=0).sum()
                                areasum.index= pd.MultiIndex.from_tuples([(area, "Total") for area in areasum.index], names=pivot.index.names)
                                frameout= []
                                for area in pivot.index.get_level_values(0).unique():
                                    baris_area= pivot.loc[area]
                                    baris_area= baris_area.reindex(urutansla, fill_value=0)
                                    baris_area.index.name= "SLA"
                                    baris_area.index= pd.MultiIndex.from_product([[area],baris_area.index], names= pivot.index.names)
                                    frameout.append(baris_area)
                                    frameout.append(areasum.loc[[area]])
                    
                                final= pd.concat(frameout)
                                totalperstatus= final.loc[(slice(None),["Total"]),:]
                                grandtotal= totalperstatus.sum(numeric_only=True).to_frame().T
                                grandtotal.index= pd.MultiIndex.from_tuples([("Grand Total", "Total")], names=final.index.names)
                                final=final.reset_index()
                                final[kol_area]= final[kol_area].where(final["SLA"].isin(["0-6 Jam", "Grand Total"]), "")
                                final= pd.concat([final,grandtotal])
                                rows_shown_amt= min(len(final),20)
                                row_height= 35
                                kolomheader=pd.MultiIndex.from_tuples([("", kol_area),("", "SLA"),("ONGOING-NOW", "OPEN"),("ONGOING-NOW", "ONPROGRESS"),("ONGOING-NOW", "POSTPONE"),("OPEN-COMPLETE", "COMPLETE"),("", "INTEGRATION FAILED"), ("", "CANCEL"), ("", "TOTAL")], names=[None, None])
                                statkolorder = [ kol_area,"SLA","OPEN","ONPROGRESS","POSTPONE","COMPLETE","INTEGRATION FAILED", "CANCEL","TOTAL"]
                                finaltabel= final[statkolorder]
                                finaltabel.columns =kolomheader
                                def warnain_baris(baris_total):
                                    text=" ".join(map(str, baris_total.values))   
                                    if "Total" in text:
                                        return ["background-color:orange "]*len(baris_total)
                                    return[""]*len(baris_total)
                                styled= finaltabel.style.apply(warnain_baris, axis=1)
                                def styletotal(dfrender):
                                    def barisnyadistyle(s):
                                        txt = " ".join(map(str, s.values))
                                        return ["background-color: orange" if "Total" in txt else ""] * len(s)
                                    return dfrender.style.apply(barisnyadistyle, axis=1)
                                def rendersla(name, dfrender, height=420):
                                    st.subheader(name)
                                    st.dataframe(styletotal(dfrender), hide_index=True, height=height)
                                
            

                                        
                                vendorfeature=(sladf.groupby([loccol, "VendorName"]).size().reset_index(name="Amount"))
                                tab_location, tab_vendor, tab_vendor2 = st.tabs(["Location", "Vendor per Location","Vendor Pivot"])
                                st.write("## \n")
                                with tab_location:
                                    col_table, col_side = st.columns([3, 1])
                                    with col_table:
                                        rendersla("Location", finaltabel, height=900)
                                    with col_side:
                                        st.write("##")
                                        if {"Region", "SubRegion", "City"}.issubset(sladf.columns):
                                            for region in sorted(sladf["Region"].dropna().unique()):
                                                regiondf= sladf[sladf["Region"]== region]
                                                region_count= len(regiondf)
                                                with st.expander(f"{region} ({region_count})"):
                                                    
                                                    for subregion in sorted(sladf["SubRegion"].dropna().unique()):
                                                        subregiondf= sladf[sladf["SubRegion"]== subregion]
                                                        subregion_count= len(subregiondf)
                                                        with st.expander(f"{subregion} ({subregion_count})"):
                                                            
                                                            citykiri, citykanan= st.columns(2)
                                                            city= sorted(subregiondf["City"].dropna().unique())
                                                            bagidua=(len(city)+1)//2
                                                            
                                                            with citykiri:
                                                                for c in city[:bagidua]:
                                                                    citydf= sladf[sladf["City"]== c]
                                                                    city_count= len(citydf)
                                                                    st.write(f"{c} ({city_count})")
                                                            with citykanan:
                                                                for c in city[bagidua:]:
                                                                    citydf= sladf[sladf["City"]== c]
                                                                    city_count= len(citydf)
                                                                    st.write(f"{c} ({city_count})")
                                        else:
                                            st.info("data ga lengkap")                           
                                                                    
                                            
                                with tab_vendor:
                                    col_table, col_side = st.columns([3, 1])
                                    with col_table:
                                        rendersla(f"Vendor per {loccol}", finaltabel, height=900)
                                    with col_side:
                                        st.write("##")
                                        for v in vendorfeature[loccol].unique():
                                            with st.expander(str(v)):
                                                subdf= vendorfeature[vendorfeature[loccol]==v]
                                                for w, row in subdf.iterrows():st.write(f"{row['VendorName']} ({row['Amount']})")
                                with tab_vendor2:
                                    
                                    vendorpivot=(tergabung.merge(sladf[["WorkOrderNumber","VendorName"]].drop_duplicates(), on="WorkOrderNumber", how="left").groupby(["VendorName","slaoptions", "StatusReport"]).agg({"WorkOrderNumber": "nunique"}).reset_index())
                                    vendorpivot= vendorpivot.pivot_table(index=["VendorName","slaoptions"], columns="StatusReport", values="WorkOrderNumber", aggfunc="sum", fill_value=0).reset_index()
                                    urutanstatus=["OPEN","ONPROGRESS","POSTPONE","COMPLETE", "INTEGRATION FAILED", "CANCEL"]
                                    urutansla = ["0-6 Jam","6-12 Jam","12-18 Jam","18-24 Jam",">24 Jam"]
                                    for ur in urutanstatus:
                                        if ur not in vendorpivot.columns:
                                            vendorpivot[ur]=0
                                    vendorpivot["TOTAL"]=vendorpivot[urutanstatus].sum(axis=1)
                                    # statkolorder = [ kol_area,"SLA","OPEN","ONPROGRESS","POSTPONE","COMPLETE","INTEGRATION FAILED", "CANCEL","TOTAL"]
                                    frame=[]
                                    for v in vendorpivot["VendorName"].unique():
                                        
                                        baris_vendor=vendorpivot[vendorpivot["VendorName"] == v].copy()
                                        baris_vendor= baris_vendor.set_index("slaoptions").reindex(urutansla, fill_value= 0).reset_index()
                                        baris_vendor["VendorName"]= v
                                        frame.append(baris_vendor)
                                        total= baris_vendor[urutanstatus+["TOTAL"]].sum().to_dict()
                                        total["slaoptions"]= "TOTAL"
                                        total["VendorName"]= v
                                        frame.append(pd.DataFrame([total]))
                                        
                                    finaldf= pd.concat(frame, ignore_index=True)
                                    finaldf["slaoptions"]= pd.Categorical(finaldf["slaoptions"],categories=urutanstatus+["TOTAL"], ordered=True)
                                    finaldf=finaldf.sort_values(["VendorName", "slaoptions"])
                                    finaldf= pd.concat(frame, ignore_index=True)
                                    finaldf["VendorName"]= finaldf["VendorName"].mask(finaldf["VendorName"].duplicated(),"")
                                    barisheader=pd.MultiIndex.from_tuples([("", "Vendor"),("", "SLA"),("ONGOING-NOW", "OPEN"),("ONGOING-NOW", "ONPROGRESS"),("ONGOING-NOW", "POSTPONE"),("OPEN-COMPLETE", "COMPLETE"),("", "INTEGRATION FAILED"), ("", "CANCEL"), ("", "TOTAL")], names=[None, None])
                                    finaldf=finaldf.reindex(columns=["VendorName", "slaoptions"]+urutanstatus+["TOTAL"])
                                    finaldf.columns=barisheader
                                    
                                    col_table, col_side = st.columns([3, 1])
                                    with col_table:
                                        st.write("## Vendor Pivot")
                                        st.dataframe(finaldf, hide_index=True, height= 900)
                                    with col_side:
                                        st.write("##")


                                # with tab_wo:
                                #     col_table, col_side = st.columns([3, 2])

                                #     with col_table:
                                        
                                #         st.subheader("Detail / Drilldown")
                                        
                                #     with col_side:
                                #         st.write("bagian kanan")


                else:
                    st.warning("data kolomny galengkap di sheet historyworkorder")
                    
            else:
                st.warning ("ganemu historywo")
                
            st.write("## ")
            st.caption("Click the button below to calculate the status duration, status SLA and export as Excel")  
            exportbutton, fillerexpbutton1, fillerexpbutton2= st.columns([1, 2, 3])
            with exportbutton:
                if st.button("Export to Excel", type="primary", use_container_width= True):
                    with st.spinner("Processing..."):
                        out, filename= exportfile(uploaded)
                    if out is not None:
                        st.success("Export ready")
                        st.download_button("Download Excel", data= out.getvalue(), file_name= filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True )
                    else:
                        st.warning("Export Failed")
    
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
