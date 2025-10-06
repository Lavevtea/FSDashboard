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
if st.sidebar.button("Data Comparation"):
    st.session_state.menu_sidebar = "Data Comparation"
if st.sidebar.button("Merge Files"):
    st.session_state.menu_sidebar = "Merge Files"
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
            'WorkOrderNumber':'WO Fieldsa',
            'WorkOrderStatusItem' : 'status',
            'Modified' : 'timestamp'  
        })

        dfcalc['status'] = dfcalc['status'].astype(str).str.strip()
        dfcalc['timestamp'] = pd.to_datetime(dfcalc['timestamp'])
        dfcalc= dfcalc.dropna(subset=['WO Fieldsa', 'timestamp'])
        dfcalc= dfcalc.sort_values(['WO Fieldsa', 'timestamp'])

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

        for wo, group in dfcalc.groupby('WO Fieldsa'):
            statusnump= group['status'].to_numpy()
            timenump= group['timestamp'].to_numpy()
            baris= {'WO Fieldsa': wo}
            timestemp={}
            lasttime=pd.Timestamp.min
            # baris['Anomali']= 'Abnormal' if pd.Series(statusnump).duplicated().any() else 'Normal'

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
                'WO Fieldsa': wo, 
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

        final=pd.DataFrame(hasil).merge(pd.DataFrame(statusdur), on='WO Fieldsa', how='left' )

        addcols= [ 'WorkOrderNumber', 'ReferenceCode', 'WorkOrderTypeName', 'DivisionName', 'WorkOrderStatusItem', 'Reason',
            'CustomerId', 'CustomerName', 'Cid', 'CircuitId', 'EndCustomerName', 'SubRegion',
            'City', 'DeviceAllocation', 'VendorName', 'DispatcherName', 'TechnicianName']

        df2= pd.read_excel(xcel, sheet_name='WorkOrder', usecols=addcols).rename(columns={'WorkOrderNumber':'WO Fieldsa'}).reset_index(drop=True)
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
            "Posted To Ax Integration Success": "COMPLETE",
            
            "Postpone": "POSTPONE",
            
            "Sms Integration Failed": "INTEGRATION FAILED",
            "Posted To Ax Integration Failed": "INTEGRATION FAILED",
            "Provisioning Failed": "INTEGRATION FAILED",
            
            "Complete With Note Request": "APPROVAL DISPATCHER FS",
            "Postpone Request": "APPROVAL DISPATCHER FS",
            
            "Cancel Work Order": "CANCEL"}
        
        final= final.merge(df2, on='WO Fieldsa', how='left')
        final['Region']= final['SubRegion'].map(regionmap).fillna('N/A')
        final['StatusReport']= final['WorkOrderStatusItem'].map(statusreportmap).fillna('N/A')
        subregionindex= final.columns.get_loc('SubRegion')
        final.insert(subregionindex, 'Region', final.pop('Region'))
        wostatusindex= final.columns.get_loc('WorkOrderStatusItem')
        final.insert(wostatusindex, 'StatusReport', final.pop('StatusReport'))

        df3= pd.read_excel(xcel, sheet_name='Rca', usecols=['WorkOrderNumber', 'UpTime']).rename(columns={'WorkOrderNumber':'WO Fieldsa'}).reset_index(drop=True)
        final= final.merge(df3, on='WO Fieldsa', how='left')
        final['UpTime'] = final['UpTime'].fillna('N/A')



        urutanstatuswo= [
            'WO Fieldsa','ReferenceCode','WorkOrderTypeName','DivisionName','CustomerId',
            'CustomerName','Cid','CircuitId','EndCustomerName','Region','SubRegion','City','DeviceAllocation',
            'VendorName','DispatcherName','TechnicianName','UpTime','Open','Assign to dispatch external','Assign to technician','Accept',
            'Travel','Arrive','On Progress','Done','Work Order Confirmation Approve','Complete',
            'Complete with note approve','Complete with note request','Complete with note reject',
            'Postpone Request','Postpone is Revised','Postpone','SMS Integration Failed','Return',
            'Return by Technician','Revise','Return is revised','Provisioning In Progress','Provisioning Success',
            'Posted to AX Integration Failed','Posted to AX Integration Success','Provisioning Failed','Cancel Work Order',
            'Open - Assign to dispatch external','SLA Open-Dispatch External',
            'Assign to dispatch external - Assign to technician','SLA Dispatch External -Technician',
            'Assign to technician - Accept','SLA Technician-Accept','Accept - Done','SLA Accept-Done',
            'Done - Complete','SLA Done-Complete','WorkOrderStatusItem','StatusReport','Reason'
        ]

        final= final[[col for col in urutanstatuswo if col in final.columns]]
        # st.write("nyanayana")
        # st.dataframe(final.head(12))
        # st.write("adsf")
        # st.write(list(final.columns))
        buffer= BytesIO()
        st.session_state.finalcopy = final.copy()
        with pd.ExcelWriter(buffer, engine= "xlsxwriter") as writer:
            final.to_excel(writer, index= False, sheet_name= "SLA")
            df_status= pd.DataFrame(list(statusreportmap.items()), columns=["Status", "Klasifikasi Status"])
            df_status.to_excel(writer, index=False, sheet_name="KeteranganStatus")
            
        buffer.seek(0)
        suggestname = f"Dashboard_FIELDSA_{pd.Timestamp.now():%Y%m%d_%H%M%S}.xlsx"
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

            ubahheader={
                "WorkOrderTypeName":"Work Order Type",
                "DivisionName": "Division",
                "Region": "Region",
                "SubRegion":"Sub Region",
                "City": "City"
            }
            
            filter1,filter2, filter3=st.columns(3)
            
            # if "locfilter" not in st.session_state:
            #     st.session_state.locfilter= "Region"
            if "divfilter" not in st.session_state:
                st.session_state.divfilter= ["Broadband", "Lms", "Fiberisasi"]
            if "tipefilter" not in st.session_state:
                st.session_state.tipefilter= ["Troubleshoot", "Activation"]
            
            
            with filter1:
                if "Region" in df.columns:
                    semuaregion= sorted(df["Region"].dropna().astype(str).unique().tolist())
                    selectedregion= st.multiselect("Select Region", options=semuaregion, default=semuaregion, key= "region_filter")
                    if selectedregion:
                        df=df[df["Region"].astype(str).isin(selectedregion)]
                if "SubRegion" in df.columns:
                    semuasubregion= sorted(df[df["Region"].astype(str).isin(selectedregion)]["SubRegion"].dropna().astype(str).unique().tolist())
                    selectedsubregion= st.multiselect("Select SubRegion", options=semuasubregion, default=semuasubregion, key= "subregion_filter")
                    if selectedregion:
                        df=df[df["SubRegion"].astype(str).isin(selectedsubregion)]
                if "City" in df.columns:
                    semuacity= sorted(df[df["SubRegion"].astype(str).isin(selectedsubregion)]["City"].dropna().astype(str).unique().tolist())
                    selectedcity= st.multiselect("Select City", options=semuacity, default=semuacity, key= "city_filter")
                    if selectedregion:
                        df=df[df["City"].astype(str).isin(selectedcity)]
                
                      
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
    
            st.divider()
            pie_wotype, pie_wodiv= st.columns(2)
            with pie_wotype:
                if "WorkOrderTypeName" in df.columns:
                    hitung_type= df["WorkOrderTypeName"].value_counts().reset_index()
                    hitung_type.columns= ["wotype", "count"]
                    tipe= px.pie(hitung_type, names="wotype", values="count",title="WO Type Chart")
                    st.plotly_chart(tipe, use_container_width=True)
                else:
                    st.warning("Tidak ada WO type yang dipilih")
            with pie_wodiv:
                if "DivisionName" in df.columns:
                    hitung_divisi= df["DivisionName"].value_counts().reset_index()
                    hitung_divisi.columns =  ["division", "count"]
                    divisi= px.pie(hitung_divisi, names="division", values="count", title="Division Chart")
                    st.plotly_chart(divisi, use_container_width=True)
                else:
                    st.warning("Tidak ada Division yang dipilih")
            st.divider()
            if not df.empty:
                for col in ["Region","SubRegion", "City"]:
                    if col in df.columns:
                        itungisi= df[col].astype(str).value_counts().reset_index()
                        itungisi.columns= [col, "Count"]
                        wo_barchart= px.bar(itungisi, x=col, y="Count", color= col, title=f"WorkOrder per {ubahheader.get(col,col)}", labels={col:ubahheader.get(col, col),"Count":"Jumlah"})
                        st.plotly_chart(wo_barchart, use_container_width= True)
                        st.dataframe(itungisi)
                        st.divider()
            else: 
                st.warning("Data kosong")
            # if loccol in df.columns:
            #     normcol = df[loccol].astype(str)
            #     itungisi = normcol.value_counts().reset_index()
            #     itungisi.columns = [loccol, "Count"]

            #     bar = px.bar(
            #         itungisi,
            #         x= loccol,
            #         y= "Count",
            #         color= loccol,
            #         title= f"Based on {ubahheader.get(loccol, loccol)}",
            #         labels= {loccol: ubahheader.get(loccol,loccol), "Count": "Jumlah"}
            #     )
            #     st.plotly_chart(bar, use_container_width=True)
            # else:
            #     st.warning("gaada kolom yang bisa dipakai")
            tambahcols = ["CustomerName", "VendorName", "Reason"]
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
            ubahheader={
                "WorkOrderTypeName":"Work Order Type",
                "DivisionName": "Division",
                "Region": "Region",
                "SubRegion":"Sub Region",
                "City": "City"
            }
            statusmap={
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
                "Posted To Ax Integration Success": "COMPLETE",
                
                "Postpone": "POSTPONE",
                
                "Sms Integration Failed": "INTEGRATION FAILED",
                "Posted To Ax Integration Failed": "INTEGRATION FAILED",
                "Provisioning Failed": "INTEGRATION FAILED",
                
                "Complete With Note Request": "APPROVAL DISPATCHER FS",
                "Postpone Request": "APPROVAL DISPATCHER FS",
                
                "Cancel Work Order": "CANCEL"}
            
            dfstat["StatusReport"]= dfstat["WorkOrderStatusItem"].map(statusmap).fillna("N/A")
            
            locfilter2,wotype2,division2,statfilter=st.columns(4)
            
            if "location2" not in st.session_state:
                st.session_state.location2= "Region"
            if "wotype2" not in st.session_state:
                st.session_state.wotype2=["Troubleshoot","Activation"]
            if "division2" not in st.session_state:
                st.session_state.division2=["Broadband","Lms", "Fiberisasi"]
            if "stat" not in st.session_state:
                st.session_state.stat= ["OPEN", "COMPLETE", "ONPROGRESS", "POSTPONE", "INTEGRATION FAILED", "APPROVAL DISPATCHER FS","CANCEL"]
            
    
            with locfilter2:
                if "Region" in dfstat.columns:
                    semuaregion= sorted(dfstat["Region"].dropna().astype(str).unique().tolist())
                    selectedregion= st.multiselect("Select Region", options=semuaregion, default=semuaregion, key= "region_filter")
                    if selectedregion:
                        dfstat=dfstat[dfstat["Region"].astype(str).isin(selectedregion)]
                if "SubRegion" in dfstat.columns:
                    semuasubregion= sorted(dfstat[dfstat["Region"].astype(str).isin(selectedregion)]["SubRegion"].dropna().astype(str).unique().tolist())
                    selectedsubregion= st.multiselect("Select SubRegion", options=semuasubregion, default=semuasubregion, key= "subregion_filter")
                    if selectedregion:
                        dfstat=dfstat[dfstat["SubRegion"].astype(str).isin(selectedsubregion)]
                if "City" in dfstat.columns:
                    semuacity= sorted(dfstat[dfstat["SubRegion"].astype(str).isin(selectedsubregion)]["City"].dropna().astype(str).unique().tolist())
                    selectedcity= st.multiselect("Select City", options=semuacity, default=semuacity, key= "city_filter")
                    if selectedregion:
                        dfstat=dfstat[dfstat["City"].astype(str).isin(selectedcity)]

            with wotype2:
                st.markdown("WO Type")
                with st.expander("WO Type"):
                    tipewo= ['Troubleshoot', 'Activation']
                    jumlahtipe= dfstat["WorkOrderTypeName"].value_counts().to_dict()
                    
                    tipeygdipilih=[]
                    
                    for r in tipewo:
                        itungg= jumlahtipe.get(r,0)
                        if st.checkbox(f"{r} ({itungg})", value= r in st.session_state.wotype2, key=f"wotype_{r}" ):
                            tipeygdipilih.append(r) 
                    st.session_state.wotype2= tipeygdipilih              
            if st.session_state.wotype2:
                dfstat= dfstat[dfstat["WorkOrderTypeName"].isin(st.session_state.wotype2)]
            else:
                st.warning("No Workorder type chosen")   
                
                       
            with division2:
                st.markdown("Division")
                with st.expander("Division"):
                    divwo= ["Broadband","Lms","Fiberisasi"]
                    jumlahperdiv= dfstat["DivisionName"].value_counts().to_dict()
                    divisiterpilih=[]
                    for c in divwo:
                        itung= jumlahperdiv.get(c,0)
                        if st.checkbox(f"{c} ({itung})", value= c in st.session_state.division2, key=f"{c}"):
                            divisiterpilih.append(c)
                    st.session_state.division2= divisiterpilih
    
            if divisiterpilih:
                dfstat= dfstat[dfstat["DivisionName"].isin(divisiterpilih)]
            else:
                st.warning("No division type chosen")
                
             

            with statfilter:
                st.markdown("Status")
                with st.expander("Status"):
                    statwo= ["OPEN", "COMPLETE", "ONPROGRESS", "POSTPONE", "INTEGRATION FAILED","APPROVAL DISPATCHER FS", "CANCEL"]
                    jumlahperstat= dfstat["StatusReport"].value_counts().to_dict()
                    
                    statusterpilih=[]
                
                    for s in statwo:
                        itung= jumlahperstat.get(s,0)
                        if st.checkbox(f"{s} ({itung})", value= (s in st.session_state.stat), key=f"status_{s}"):
                            statusterpilih.append(s)
                    st.session_state.stat= statusterpilih
                # statwo= ["OPEN", "COMPLETE", "ONPROGRESS", "POSTPONE", "INTEGRATION FAILED","APPROVAL DISPATCHER FS", "CANCEL"]
                # jumlahperstat= dfstat["StatusReport"].value_counts().to_dict()
                # allstat= [a for a in statwo if a in dfstat["StatusReport"].unique()]
                # opsi= [f"{a} ({jumlahperstat.get(a, 0)})" for a in allstat]
                # default= opsi
                # statusterpilih = st.checkbox("Select Status", options= opsi, default= default, key= "status_filter")
                    
            if statusterpilih:
                dfstat= dfstat[dfstat["StatusReport"].isin(statusterpilih)]
            else:
                st.warning("No status picked")
            
            if not dfstat.empty:
                st.divider()
                pie_wotipe, pie_divisi, pie_statusreport =  st.columns(3)
                with pie_wotipe:
                    if "WorkOrderTypeName" in dfstat.columns:
                        tipe_count= dfstat["WorkOrderTypeName"].value_counts().reset_index()
                        tipe_count.columns= ["wotipe", "count"]
                        pie_tipe= px.pie(tipe_count, names= "wotipe", values= "count", title= "WO Type Chart")
                        st.plotly_chart(pie_tipe, use_container_width= True)
                    else:
                        st.warning("Tidak ada WO Type yang dipilih")
                with pie_divisi:
                    if "DivisionName" in dfstat.columns:
                        tipe_div= dfstat["DivisionName"].value_counts().reset_index()
                        tipe_div.columns= ["divisi", "count"]
                        pie_div= px.pie(tipe_div, names= "divisi", values= "count", title= "Division Chart")
                        st.plotly_chart(pie_div, use_container_width= True)
                    else:
                        st.warning("Tidak ada Division yang dipilih")
                with pie_statusreport:
                    if "StatusReport" in dfstat.columns:
                        stat_count= dfstat["StatusReport"].value_counts().reset_index()
                        stat_count.columns= ["stat", "count"]
                        pie_stat= px.pie(stat_count, names= "stat", values= "count", title= "StatusReport Chart")
                        st.plotly_chart(pie_stat, use_container_width= True)
                    else:
                        st.warning("Tidak ada Status yang dipilih")


            for s in dfstat.columns:
                if dfstat[s].dtype== object:
                    dfstat[s]= dfstat[s].astype(str).str.strip()
            if "Region" not in dfstat.columns:
                dfstat["Region"]= dfstat["SubRegion"].map(regionmap).fillna(dfstat["SubRegion"])
        
            if tipeygdipilih:
                if (isinstance(exceldata, dict)and "HistoryWorkOrder" in exceldata):
                    stathistory= exceldata["HistoryWorkOrder"].copy()
                    stathistory.columns=stathistory.columns.astype(str).str.strip()
                    his_wo= "WorkOrderNumber" if "WorkOrderNumber" in stathistory.columns else None
                    his_stat= "WorkOrderStatusItem" if "WorkOrderStatusItem" in stathistory.columns else None
                    his_time= "Modified" if "Modified" in stathistory.columns else None
                    if all([his_wo, his_stat, his_time]):
                        scannedwo=(dfstat[["WorkOrderNumber","WorkOrderStatusItem"]].dropna().copy())
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
                                if a in statusmap:
                                    return statusmap[a]
                                return statusmap.get(str(a).title(), "OTHER")
                            gabungsemua["StatusReport"]=  gabungsemua["StatusWO"].apply(statusreportmap)
                            gabungsemua= gabungsemua[gabungsemua["StatusReport"].isin(["OPEN","ONPROGRESS","POSTPONE","COMPLETE", "INTEGRATION FAILED","APPROVAL DISPATCHER FS", "CANCEL"])].copy()
                            gabungsemua["wonum_key"]= gabungsemua["WorkOrderNumber"].astype(str).str.strip().str.title()                     
                            gabungsemua["statusnormalized"]= gabungsemua["StatusWO"].astype(str).fillna("").str.strip().str.title()
                            tergabung= gabungsemua.merge(scannedwo[["wonum_key", "statusnormalized"]], on="wonum_key", how="inner", suffixes=("", "wo"))
                            tergabung=tergabung[tergabung["statusnormalized"]==tergabung["statusnormalizedwo"]].copy()

                            if tergabung.empty:
                                st.write("gaada yg match antara history stat dan stat excel")       
                            else:
                                tergabung["WorkOrderNumber"]= tergabung["WorkOrderNumber"].astype(str).str.strip().str.upper()
                                dfstat["WorkOrderNumber"]= dfstat["WorkOrderNumber"].astype(str).str.strip().str.upper()
                                tergabung= tergabung.merge(dfstat[["WorkOrderNumber","uptime"]], on="WorkOrderNumber", how="left")
                                tergabung["duration"]=((tergabung["uptime"]-tergabung["terpilih_c"]).dt.total_seconds()/3600)
                                completeonly= tergabung["StatusReport"]=="COMPLETE"
                                tergabung.loc[completeonly, "duration"]= ((tergabung.loc[completeonly, "status_c"]-tergabung.loc[completeonly, "open_c"]).dt.total_seconds()/3600)
                                # st.write("Distribusi Duration (jam):", tergabung["duration"].describe())
                                # st.write(tergabung[["WorkOrderNumber","StatusReport","duration"]].head(20))

                                def slaoptions_general(hour):
                                    if hour<=4:
                                        return "0-4 Jam"
                                    elif hour<=6:
                                        return "4-6 Jam"
                                    elif hour<=12:
                                        return "6-12 Jam"
                                    elif pd.isna(hour):
                                        return None
                                    else:
                                        return ">12 Jam"
                                def slaoptions_broadband(hour):
                                    if hour<=6:
                                        return "0-6 Jam"
                                    elif hour<=12:
                                        return "6-12 Jam"
                                    elif hour<=24:
                                        return "12-24 Jam"
                                    else:
                                        return ">24 Jam"
                                broadband_only= (len(st.session_state.division2)==1 and "Broadband" in st.session_state.division2)
                                if broadband_only:
                                    tergabung["slaoptions"]= tergabung["duration"].apply(slaoptions_broadband)
                                    urutansla= ["0-6 Jam", "6-12 Jam", "12-24 Jam", ">24 Jam"]
                                else:
                                    tergabung["slaoptions"]= tergabung["duration"].apply(slaoptions_general)
                                    urutansla= ["0-4 Jam", "4-6 Jam", "6-12 Jam", ">12 Jam"]
                                priority_cols = [
                                    ("Region", selectedregion),
                                    ("SubRegion", selectedsubregion),
                                    ("City", selectedcity)
                                ]

                                kol_area = None
                                for col, selected in priority_cols:
                                    if col in dfstat.columns and len(selected) > 0:
                                        kol_area = col
                                        break  
                                if kol_area is None:
                                    kol_area = "City"

                                tergabung= tergabung.merge(dfstat[["WorkOrderNumber", "uptime", "City", "SubRegion", "Region"]], on="WorkOrderNumber", how="left")
                                # st.write(tergabung)
                                tergabung_valid= tergabung.dropna(subset=["slaoptions"]).copy()
                                # st.write(tergabung_valid)

                                
                                slagroup=( tergabung_valid.groupby([kol_area, "StatusReport", "slaoptions"]).agg({"WorkOrderNumber": "nunique"}).reset_index())
                                                                                                                
                                if "WorkOrderNumber" in slagroup.columns:
                                    slagroup=slagroup.rename(columns={"WorkOrderNumber":"Count"})
                                slagroup["Count"]= slagroup["Count"].astype(int)
                                urutanstatus=["OPEN","ONPROGRESS","POSTPONE","COMPLETE", "INTEGRATION FAILED", "APPROVAL DISPATCHER FS", "CANCEL"]
                                def buat_sla_table(tergabung_valid, kol_area, urutansla, urutanstatus):
                                    slagroup = (
                                        tergabung_valid.groupby([kol_area, "StatusReport", "slaoptions"])
                                        .agg({"WorkOrderNumber": "nunique"})
                                        .reset_index()
                                    )

                                    if "WorkOrderNumber" in slagroup.columns:
                                        slagroup = slagroup.rename(columns={"WorkOrderNumber":"Count"})
                                    slagroup["Count"]= slagroup["Count"].astype(int)

                                    pivot = slagroup.pivot_table(
                                        index=[kol_area,"slaoptions"],
                                        columns="StatusReport",
                                        values="Count",
                                        aggfunc="sum",
                                        fill_value=0
                                    )
                                    pivot.index.set_names([kol_area, "SLA"], inplace=True)
                                    area_list= tergabung_valid[kol_area].dropna().unique().tolist()
                                    full_index= pd.MultiIndex.from_product([area_list, urutansla], names= [kol_area, "SLA"])
                                    pivot= pivot.reindex(full_index, fill_value=0)
                                    # st.write(area_list)
                                    for u in urutanstatus:
                                        if u not in pivot.columns:
                                            pivot[u]=0

                                    pivot = pivot.reindex(columns=urutanstatus, fill_value=0)
                                    pivot["TOTAL2"]= pivot[["OPEN","ONPROGRESS","POSTPONE"]].sum(axis=1)
                                    pivot["TOTAL"]= pivot[["OPEN","ONPROGRESS","POSTPONE","COMPLETE","INTEGRATION FAILED","APPROVAL DISPATCHER FS","CANCEL"]].sum(axis=1)

                                    pivot.index.set_names([kol_area, "SLA"], inplace=True)

                                    areasum= pivot.groupby(level=0).sum()
                                    areasum.index= pd.MultiIndex.from_tuples(
                                        [(area,"Total") for area in areasum.index],
                                        names=pivot.index.names
                                    )

                                    frameout=[]
                                    for area in pivot.index.get_level_values(0).unique():
                                        baris_area= pivot.loc[area]
                                        baris_area= baris_area.reindex(urutansla, fill_value=0)
                                        baris_area.index.name="SLA"
                                        baris_area.index= pd.MultiIndex.from_product([[area], baris_area.index], names=pivot.index.names)
                                        frameout.append(baris_area)
                                        frameout.append(areasum.loc[[area]])

                                    final= pd.concat(frameout)
                                    totalperstatus= final.loc[(slice(None),["Total"]),:]
                                    grandtotal= totalperstatus.sum(numeric_only=True).to_frame().T
                                    grandtotal[kol_area]= "Grand Total"
                                    grandtotal["SLA"]= ""
                                    grandtotal= grandtotal[final.reset_index().columns]
                                    final= final.reset_index()
                                    final[kol_area]= final[kol_area].where(final["SLA"].isin([urutansla[0],"Grand Total"]), "")
                                    final= pd.concat([final, grandtotal], ignore_index=True)

                                    statkolorder=[kol_area,"SLA","OPEN","ONPROGRESS","POSTPONE","TOTAL2","COMPLETE","INTEGRATION FAILED","APPROVAL DISPATCHER FS","CANCEL","TOTAL"]
                                    finaltabel= final[statkolorder]

                                    kolomheader=pd.MultiIndex.from_tuples([
                                        ("", kol_area),("", "SLA"),
                                        ("ONGOING-NOW","OPEN"),("ONGOING-NOW","ONPROGRESS"),("ONGOING-NOW","POSTPONE"),("ONGOING-NOW","TOTAL"),
                                        ("OPEN-COMPLETE","COMPLETE"),("", "INTEGRATION FAILED"),
                                        ("COMP NOTE REQ & POSTPONE REQ","APPROVAL DISPATCHER FS"),("", "CANCEL"),("", "TOTAL")
                                    ], names=[None,None])
                                    finaltabel.columns= kolomheader

                                    return finaltabel
                                
                                finaltabel_region = buat_sla_table(tergabung_valid, "Region", urutansla, urutanstatus)
                                finaltabel_subregion = buat_sla_table(tergabung_valid, "SubRegion", urutansla, urutanstatus)
                                finaltabel_city = buat_sla_table(tergabung_valid, "City", urutansla, urutanstatus)
                                def styletotal(dfrender):
                                    def warnain_baris(total):
                                        text=" ".join(map(str, total.values))   
                                        return ["background-color:rgb(240, 242, 246 " if "Total" in text else ""]*len(total)
                                  
                                    def warnain_kolom(total):
                                        return ["background-color: rgb(240, 242, 246)" if total.name[1] in ["TOTAL2","TOTAL"] else ""]*len(total)
                                    
                                    def warnain_approval(col):
                                        if col.name[1] == "APPROVAL DISPATCHER FS":
                                            return ["background-color: rgb(252, 236, 3)" if p>0 else "" for p in col.values]
                                        else:
                                            return["" for g in col]
                                    
                                    def warnain_ongoing(row):
                                        styles = [""] * len(row)  
                                        sla = row[("", "SLA")]
                                        divisi= st.session_state.get("division2", [])

                                        for i, colname in enumerate(row.index):
                                            if colname[0] == "ONGOING-NOW" and colname[1] in ["OPEN", "ONPROGRESS", "POSTPONE"]:
                                                val = row[colname]
                                                if val>0:
                                                    if any(d in ["Lms", "Fiberisasi"] for d in divisi):
                                                        if sla == "4-6 Jam":
                                                            if colname[1] == "OPEN":
                                                                styles[i] = "background-color: rgb(252, 236, 3)"
                                                            elif colname[1] == "ONPROGRESS":
                                                                styles[i] = "background-color: rgb(252, 236, 3)"
                                                            elif colname[1] == "POSTPONE":
                                                                styles[i] = "background-color: rgb(252, 236, 3)"

                                                        elif sla == "6-12 Jam":
                                                            if colname[1] == "OPEN":
                                                                styles[i] = "background-color: orange"
                                                            elif colname[1] == "ONPROGRESS":
                                                                styles[i] = "background-color: orange"
                                                            elif colname[1] == "POSTPONE":
                                                                styles[i] = "background-color: orange"
                                                                
                                                        elif sla == ">12 Jam":
                                                            if colname[1] == "OPEN":
                                                                styles[i] = "background-color: red"
                                                            elif colname[1] == "ONPROGRESS":
                                                                styles[i] = "background-color: red"
                                                            elif colname[1] == "POSTPONE":
                                                                styles[i] = "background-color: red"
                                                    elif "Broadband" in divisi:
                                                        if sla == "12-24 Jam":
                                                            if colname[1] == "OPEN":
                                                                styles[i] = "background-color: rgb(252, 236, 3)"
                                                            elif colname[1] == "ONPROGRESS":
                                                                styles[i] = "background-color: rgb(252, 236, 3)"
                                                            elif colname[1] == "POSTPONE":
                                                                styles[i] = "background-color: rgb(252, 236, 3)"

                                                        elif sla == ">24 Jam":
                                                            if colname[1] == "OPEN":
                                                                styles[i] = "background-color: red"
                                                            elif colname[1] == "ONPROGRESS":
                                                                styles[i] = "background-color: red"
                                                            elif colname[1] == "POSTPONE":
                                                                styles[i] = "background-color: red"

                                        return styles

                                        # if col.name[0] == "ONGOING-NOW":
                                        #     warna= []
                                        #     sla_values= dfrender.index.get_level_values("SLA")
                                        #     for index, val in enumerate(col.values):
                                        #         sla= sla_values[index]
                                        #         if any(div in ["Lms", "Fiberisasi"] for div in st.session_state.division2 ):
                                        #             if sla == "4-6 Jam" and val>0:
                                        #                 warna.append("background-color: red")
                                                
                                        #             elif sla == "6-12 Jam" and val>0:
                                        #                 warna.append("background-color: orange")
                                        #             else:
                                        #                 warna.append("")
                                        #         else:
                                        #                 warna.append("")
                                        #     return warna
                                        # else:
                                        #     return ["" for c in col]
                                            
                                    return(dfrender.style.apply(warnain_approval, axis=0).apply(warnain_ongoing, axis=1).apply(warnain_kolom, axis=0).apply(warnain_baris, axis=1))
                                    
                                def rendersla(name, dfrender, kol_area=None, selected=None, height=420):
                                    if kol_area and selected:
                                        dfrender = dfrender[dfrender.iloc[:,0].isin(selected)]  
                                    st.subheader(name)
                                    st.dataframe(styletotal(dfrender), hide_index=True, height=height)
                            
                                regionbar= (tergabung_valid.groupby(["Region", "StatusReport"]).size().reset_index(name="Count"))
                                chart_region= px.bar(regionbar, x="Region", y="Count", color="StatusReport", barmode="stack", title="Status Chart per Region", labels={"Region":"Region","Count":"Jumlah"})
                                
                                
                                subregionbar= (tergabung_valid.groupby(["SubRegion", "StatusReport"]).size().reset_index(name="Count"))
                                chart_subregion= px.bar(subregionbar, x="SubRegion", y="Count", color="StatusReport", barmode="stack", title="Status Chart per Sub Region", labels={"SubRegion":"Sub Region","Count":"Jumlah"})
                                
                                
                                citybar= (tergabung_valid.groupby(["City", "StatusReport"]).size().reset_index(name="Count"))
                                chart_city= px.bar(citybar, x="City", y="Count", color="StatusReport", barmode="stack", title="Status Chart per City", labels={"City":"City","Count":"Jumlah"})
                                
                                st.plotly_chart(chart_region, use_container_width= True)
                                rendersla("SLA Summary per Region", finaltabel_region, kol_area="Region",height=500)
                                
                                st.plotly_chart(chart_subregion, use_container_width= True)
                                rendersla("SLA Summary per SubRegion", finaltabel_subregion, kol_area="SubRegion", height=500)
                                
                                st.plotly_chart(chart_city, use_container_width= True)
                                rendersla("SLA Summary per City", finaltabel_city, kol_area="City", height=500)
                                
                            
                            
                else: 
                    st.warning("Data kosong")
            

                # st.dataframe(statussummary)
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
            "Posted To Ax Integration Success": "COMPLETE",
            
            "Postpone": "POSTPONE",
            
            "Sms Integration Failed": "INTEGRATION FAILED",
            "Posted To Ax Integration Failed": "INTEGRATION FAILED",
            "Provisioning Failed": "INTEGRATION FAILED",
            
            "Complete With Note Request": "APPROVAL DISPATCHER FS",
            "Postpone Request": "APPROVAL DISPATCHER FS",
            
            "Cancel Work Order": "CANCEL"}

          
            sladf= df.copy()
            sladf["WorkOrderStatusItem"] = sladf["WorkOrderStatusItem"].astype(str).str.strip().str.title()
            sladf["StatusReport"]= sladf["WorkOrderStatusItem"].map(statusmap1).fillna("N/A")
            
            

            
            kolsla1, kolsla2, kolsla3,kolsla4 = st.columns(4)
            if "kolsla1" not in st.session_state:
                st.session_state.kolsla1= "Region"
            if "kolsla2" not in st.session_state:
                st.session_state.kolsla2=["Troubleshoot","Activation"]
            if "kolsla3" not in st.session_state:
                st.session_state.kolsla3=["Broadband","Lms", "Fiberisasi"]
            if "kolsla4" not in st.session_state:
                st.session_state.kolsla4= ["OPEN", "COMPLETE", "ONPROGRESS", "POSTPONE", "INTEGRATION FAILED", "APPROVAL DISPATCHER FS","CANCEL"]
                
            with kolsla1:
                if "Region" in sladf.columns:
                    semuaregion= sorted(sladf["Region"].dropna().astype(str).unique().tolist())
                    selectedregion= st.multiselect("Select Region", options=semuaregion, default=semuaregion, key= "region_fil")
                    if selectedregion:
                        sladf=sladf[sladf["Region"].astype(str).isin(selectedregion)]
                if "SubRegion" in sladf.columns:
                    semuasubregion= sorted(sladf[sladf["Region"].astype(str).isin(selectedregion)]["SubRegion"].dropna().astype(str).unique().tolist())
                    selectedsubregion= st.multiselect("Select SubRegion", options=semuasubregion, default=semuasubregion, key= "subregion_fil")
                    if selectedregion:
                        sladf=sladf[sladf["SubRegion"].astype(str).isin(selectedsubregion)]
                if "City" in sladf.columns:
                    semuacity= sorted(sladf[sladf["SubRegion"].astype(str).isin(selectedsubregion)]["City"].dropna().astype(str).unique().tolist())
                    selectedcity= st.multiselect("Select City", options=semuacity, default=semuacity, key= "city_fil")
                    if selectedregion:
                        sladf=sladf[sladf["City"].astype(str).isin(selectedcity)]
                        
            with kolsla2:
                st.markdown("WO Type")
                with st.expander("WO Type"):
                    wotype= ['Troubleshoot', 'Activation']
                    jumlahtype= sladf["WorkOrderTypeName"].value_counts().to_dict()
                    type_terpilih=[]
                    for w in wotype:
                        count= jumlahtype.get(w,0)
                        if st.checkbox(f"{w}({count})", value=w in st.session_state.kolsla2, key=f"wotypesla_{w}"):
                            type_terpilih.append(w)
                    st.session_state.kolsla2=type_terpilih
            if st.session_state.kolsla2:
                sladf= sladf[sladf["WorkOrderTypeName"].isin(st.session_state.kolsla2)]
            else:
                st.warning("No Workorder type chosen")
                   
            with kolsla3:
                st.markdown("Division")
                with st.expander("Division"):
                    divwo= ["Broadband","Lms","Fiberisasi"]
                    jumlahperdiv= sladf["DivisionName"].value_counts().to_dict()
                    div_terpilih=[]
                    for c in divwo:
                        itung= jumlahperdiv.get(c,0)
                        if st.checkbox(f"{c} ({itung})", value= c in st.session_state.kolsla3, key=f"{c}_sla"):
                            div_terpilih.append(c)
                    st.session_state.kolsla3= div_terpilih
            if div_terpilih:
                sladf= sladf[sladf["DivisionName"].isin(div_terpilih)]
            else:
                st.warning("No Division type chosen")
                
                    
            with kolsla4:
                st.markdown("Status")
                with st.expander("Status"):
                    statwo= ["OPEN", "COMPLETE", "ONPROGRESS", "POSTPONE", "INTEGRATION FAILED","APPROVAL DISPATCHER FS", "CANCEL"]
                    jumlahperstat= sladf["StatusReport"].value_counts().to_dict()
                    
                    status_terpilih=[]
                
                    for s in statwo:
                        itung= jumlahperstat.get(s,0)
                        if st.checkbox(f"{s} ({itung})", value= (s in st.session_state.kolsla4), key=f"status_{s}"):
                            status_terpilih.append(s)
                    st.session_state.kolsla4= status_terpilih
              
            if status_terpilih:
                sladf= sladf[sladf["StatusReport"].isin(status_terpilih)]
            else:
                st.warning("No status picked")
        
                    
    
            for a in sladf.columns:
                if sladf[a].dtype== object:
                    sladf[a]= sladf[a].astype(str).str.strip()
            if "Region" not in sladf.columns:
                sladf["Region"]= sladf["SubRegion"].map(regionmap).fillna(sladf["SubRegion"])
            


            for s in sladf.columns:
                if sladf[s].dtype== object:
                    sladf[s]= sladf[s].astype(str).str.strip()
            if "Region" not in sladf.columns:
                sladf["Region"]= sladf["SubRegion"].map(regionmap).fillna(sladf["SubRegion"])
                
                        
                        
            if type_terpilih:          
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
                            gabungsemua= gabungsemua[gabungsemua["StatusReport"].isin(["OPEN","ONPROGRESS","POSTPONE","COMPLETE", "INTEGRATION FAILED","APPROVAL DISPATCHER FS", "CANCEL"])].copy()
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
                                def slaoptions_general(hour):
                                    if hour<=4:
                                        return "0-4 Jam"
                                    elif hour<=6:
                                        return "4-6 Jam"
                                    elif hour<=12:
                                        return "6-12 Jam"
                                    elif pd.isna(hour):
                                        return None
                                    else:
                                        return ">12 Jam"
                                def slaoptions_broadband(hour):
                                    if hour<=6:
                                        return "0-6 Jam"
                                    elif hour<=12:
                                        return "6-12 Jam"
                                    elif hour<=24:
                                        return "12-24 Jam"
                                    else:
                                        return ">24 Jam"
                                broadband_only= (len(st.session_state.kolsla3)==1 and "Broadband" in st.session_state.kolsla3)
                                if broadband_only:
                                    tergabung["slaoptions"]= tergabung["duration"].apply(slaoptions_broadband)
                                    urutansla= ["0-6 Jam", "6-12 Jam", "12-24 Jam", ">24 Jam"]
                                else:
                                    tergabung["slaoptions"]= tergabung["duration"].apply(slaoptions_general)
                                    urutansla= ["0-4 Jam", "4-6 Jam", "6-12 Jam", ">12 Jam"]
                                priority_cols = [
                                    ("Region", selectedregion),
                                    ("SubRegion", selectedsubregion),
                                    ("City", selectedcity)
                                ]

                                kol_area = None
                                for col, selected in priority_cols:
                                    if col in sladf.columns and len(selected) > 0:
                                        kol_area = col
                                        break  
                                if kol_area is None:
                                    kol_area = "City"

                                tergabung= tergabung.merge(sladf[["WorkOrderNumber", "uptime", "City", "SubRegion", "Region"]], on="WorkOrderNumber", how="left")
                                # st.write(tergabung)
                                tergabung_valid= tergabung.dropna(subset=["slaoptions"]).copy()
                                # st.write(tergabung_valid)

                                
                                slagroup=( tergabung_valid.groupby([kol_area, "StatusReport", "slaoptions"]).agg({"WorkOrderNumber": "nunique"}).reset_index())
                                                                                                                
                                
                                urutanstatus=["OPEN","ONPROGRESS","POSTPONE","COMPLETE", "INTEGRATION FAILED", "APPROVAL DISPATCHER FS", "CANCEL"]
                                def buat_sla_table(tergabung_valid, kol_area, urutansla, urutanstatus):
                                    slagroup = (
                                        tergabung_valid.groupby([kol_area, "StatusReport", "slaoptions"])
                                        .agg({"WorkOrderNumber": "nunique"})
                                        .reset_index()
                                    )
                                    if "WorkOrderNumber" in slagroup.columns:
                                        slagroup=slagroup.rename(columns={"WorkOrderNumber":"Count"})
                                    slagroup["Count"]= slagroup["Count"].astype(int)
                                    pivot = slagroup.pivot_table(
                                        index=[kol_area,"slaoptions"],
                                        columns="StatusReport",
                                        values="Count",
                                        aggfunc="sum",
                                        fill_value=0
                                    )
                                    pivot.index.set_names([kol_area, "SLA"], inplace=True)
                                    area_list= tergabung_valid[kol_area].dropna().unique().tolist()
                                    full_index= pd.MultiIndex.from_product([area_list, urutansla], names= [kol_area, "SLA"])
                                    pivot= pivot.reindex(full_index, fill_value=0)
                                    # st.write(area_list)
                                    for u in urutanstatus:
                                        if u not in pivot.columns:
                                            pivot[u]=0

                                    pivot = pivot.reindex(columns=urutanstatus, fill_value=0)
                                    pivot["TOTAL2"]= pivot[["OPEN","ONPROGRESS","POSTPONE"]].sum(axis=1)
                                    pivot["TOTAL"]= pivot[["OPEN","ONPROGRESS","POSTPONE","COMPLETE","INTEGRATION FAILED","APPROVAL DISPATCHER FS","CANCEL"]].sum(axis=1)

                                    pivot.index.set_names([kol_area, "SLA"], inplace=True)

                                    areasum= pivot.groupby(level=0).sum()
                                    areasum.index= pd.MultiIndex.from_tuples(
                                        [(area,"Total") for area in areasum.index],
                                        names=pivot.index.names
                                    )

                                    frameout=[]
                                    for area in pivot.index.get_level_values(0).unique():
                                        baris_area= pivot.loc[area]
                                        baris_area= baris_area.reindex(urutansla, fill_value=0)
                                        baris_area.index.name="SLA"
                                        baris_area.index= pd.MultiIndex.from_product([[area], baris_area.index], names=pivot.index.names)
                                        frameout.append(baris_area)
                                        frameout.append(areasum.loc[[area]])

                                    final= pd.concat(frameout)
                                    totalperstatus= final.loc[(slice(None),["Total"]),:]
                                    grandtotal= totalperstatus.sum(numeric_only=True).to_frame().T
                                    grandtotal[kol_area]= "Grand Total"
                                    grandtotal["SLA"]= ""
                                    grandtotal= grandtotal[final.reset_index().columns]
                                    final= final.reset_index()
                                    final[kol_area]= final[kol_area].where(final["SLA"].isin([urutansla[0],"Grand Total"]), "")
                                    final= pd.concat([final, grandtotal], ignore_index=True)

                                    statkolorder=[kol_area,"SLA","OPEN","ONPROGRESS","POSTPONE","TOTAL2","COMPLETE","INTEGRATION FAILED","APPROVAL DISPATCHER FS","CANCEL","TOTAL"]
                                    finaltabel= final[statkolorder]

                                    kolomheader=pd.MultiIndex.from_tuples([
                                        ("", kol_area),("", "SLA"),
                                        ("ONGOING-NOW","OPEN"),("ONGOING-NOW","ONPROGRESS"),("ONGOING-NOW","POSTPONE"),("ONGOING-NOW","TOTAL"),
                                        ("OPEN-COMPLETE","COMPLETE"),("", "INTEGRATION FAILED"),
                                        ("COMP NOTE REQ & POSTPONE REQ","APPROVAL DISPATCHER FS"),("", "CANCEL"),("", "TOTAL")
                                    ], names=[None,None])
                                    finaltabel.columns= kolomheader

                                    return finaltabel
                            
                                finaltabel_region = buat_sla_table(tergabung_valid, "Region", urutansla, urutanstatus)
                                finaltabel_subregion = buat_sla_table(tergabung_valid, "SubRegion", urutansla, urutanstatus)
                                finaltabel_city = buat_sla_table(tergabung_valid, "City", urutansla, urutanstatus)
                                
                                def styletotal(dfrender):
                                    def warnain_baris(total):
                                        text=" ".join(map(str, total.values))   
                                        return ["background-color:rgb(240, 242, 246 " if "Total" in text else ""]*len(total)
                                    
                                    def warnain_kolom(total):
                                        return ["background-color: rgb(240, 242, 246)" if total.name[1] in ["TOTAL2","TOTAL"] else ""]*len(total)
                                    
                                    def warnain_approval(col):
                                        if col.name[1] == "APPROVAL DISPATCHER FS":
                                            return ["background-color: rgb(252, 236, 3)" if p>0 else "" for p in col.values]
                                        else:
                                            return["" for g in col]
                                    def warnain_ongoing(row):
                                        styles = [""] * len(row)  
                                        sla = row[("", "SLA")]
                                        divisi= st.session_state.get("kolsla3", [])

                                        for i, colname in enumerate(row.index):
                                            if colname[0] == "ONGOING-NOW" and colname[1] in ["OPEN", "ONPROGRESS", "POSTPONE"]:
                                                val = row[colname]
                                                if val>0:
                                                    if any(d in ["Lms", "Fiberisasi"] for d in divisi):
                                                        if sla == "4-6 Jam":
                                                            if colname[1] == "OPEN":
                                                                styles[i] = "background-color: rgb(252, 236, 3)"
                                                            elif colname[1] == "ONPROGRESS":
                                                                styles[i] = "background-color: rgb(252, 236, 3)"
                                                            elif colname[1] == "POSTPONE":
                                                                styles[i] = "background-color: rgb(252, 236, 3)"

                                                        elif sla == "6-12 Jam":
                                                            if colname[1] == "OPEN":
                                                                styles[i] = "background-color: orange"
                                                            elif colname[1] == "ONPROGRESS":
                                                                styles[i] = "background-color: orange"
                                                            elif colname[1] == "POSTPONE":
                                                                styles[i] = "background-color: orange"
                                                                
                                                        elif sla == ">12 Jam":
                                                            if colname[1] == "OPEN":
                                                                styles[i] = "background-color: red"
                                                            elif colname[1] == "ONPROGRESS":
                                                                styles[i] = "background-color: red"
                                                            elif colname[1] == "POSTPONE":
                                                                styles[i] = "background-color: red"
                                                    elif "Broadband" in divisi:
                                                        if sla == "12-24 Jam":
                                                            if colname[1] == "OPEN":
                                                                styles[i] = "background-color: rgb(252, 236, 3)"
                                                            elif colname[1] == "ONPROGRESS":
                                                                styles[i] = "background-color: rgb(252, 236, 3)"
                                                            elif colname[1] == "POSTPONE":
                                                                styles[i] = "background-color: rgb(252, 236, 3)"

                                                        elif sla == ">24 Jam":
                                                            if colname[1] == "OPEN":
                                                                styles[i] = "background-color: red"
                                                            elif colname[1] == "ONPROGRESS":
                                                                styles[i] = "background-color: red"
                                                            elif colname[1] == "POSTPONE":
                                                                styles[i] = "background-color: red"

                                        return styles
                                        
                                    return(dfrender.style.apply(warnain_approval, axis=0).apply(warnain_ongoing, axis=1).apply(warnain_kolom, axis=0).apply(warnain_baris, axis=1))
                                    
                                def rendersla(name, dfrender, kol_area=None, selected=None, height=420):
                                    if kol_area and selected:
                                        dfrender = dfrender[dfrender.iloc[:,0].isin(selected)]  # kolom area = kolom pertama
                                    st.subheader(name)
                                    st.dataframe(styletotal(dfrender), hide_index=True, height=height)
                                    # warna_bar= {"OPEN":FF2305,
                                    #             "ONPROGRESS":"FFAF1A",
                                    #             "POSTPONE": "FFAF1A",
                                    #             "COMPLETE": "009E41",
                                    #             "INTEGRATION FAILED",
                                    #             "APPROVAL DISPATCHER FS",
                                    #             "CANCEL"}
                               
                                
                                region, subregion, city,  vendor2= st.tabs(["Region","Sub Region", "City",  "Vendor Pivot"])
                                with region:
                                    rendersla("SLA Summary per Region", finaltabel_region, kol_area="Region",height=500)
                                with subregion:
                                    rendersla("SLA Summary per SubRegion", finaltabel_subregion, kol_area="SubRegion", height=500)
                                with city:
                                    rendersla("SLA Summary per City", finaltabel_city, kol_area="City", height=500) 
                            
                                            
                                # with vendor1:
                                #     col_table, col_side = st.columns([3, 1])
                                #     with col_table:
                                #         rendersla(f"Vendor per {loccol}", finaltabel, height=900)
                                #     with col_side:
                                #         st.write("##")
                                #         for v in vendorfeature[loccol].unique():
                                #             with st.expander(str(v)):
                                #                 subdf= vendorfeature[vendorfeature[loccol]==v]
                                #                 for w, row in subdf.iterrows():st.write(f"{row['VendorName']} ({row['Amount']})")
                                                
                                with vendor2:
                                    st.write("## Vendor Pivot")
                                    vendorpivot=(tergabung.merge(sladf[["WorkOrderNumber","VendorName"]].drop_duplicates(), on="WorkOrderNumber", how="left").groupby(["VendorName","slaoptions", "StatusReport"]).agg({"WorkOrderNumber": "nunique"}).reset_index())
                                    vendorpivot= vendorpivot.pivot_table(index=["VendorName","slaoptions"], columns="StatusReport", values="WorkOrderNumber", aggfunc="sum", fill_value=0).reset_index()
                                    urutanstatus=["OPEN","ONPROGRESS","POSTPONE","COMPLETE", "INTEGRATION FAILED", "APPROVAL DISPATCHER FS", "CANCEL"]
                                    for ur in urutanstatus:
                                        if ur not in vendorpivot.columns:
                                            vendorpivot[ur]=0
                                    vendorpivot["TOTAL2"]=vendorpivot[["OPEN", "ONPROGRESS", "POSTPONE"]].sum(axis=1)
                                    vendorpivot["Total"]=vendorpivot[["OPEN", "ONPROGRESS", "POSTPONE","COMPLETE", "INTEGRATION FAILED","APPROVAL DISPATCHER FS", "CANCEL"]].sum(axis=1)
                                    # statkolorder = [ kol_area,"SLA","OPEN","ONPROGRESS","POSTPONE","COMPLETE","INTEGRATION FAILED", "CANCEL","TOTAL"]
                                    frame=[]
                                    for v in vendorpivot["VendorName"].unique():
                                        
                                        baris_vendor=vendorpivot[vendorpivot["VendorName"] == v].copy()
                                        baris_vendor= baris_vendor.set_index("slaoptions").reindex(urutansla, fill_value= 0).reset_index()
                                        baris_vendor["VendorName"]= v
                                        frame.append(baris_vendor)
                                        total= baris_vendor[urutanstatus+["TOTAL2","Total"]].sum().to_dict()
                                        total["slaoptions"]= "Total"
                                        total["VendorName"]= v
                                        frame.append(pd.DataFrame([total]))
                                        
                                    finaldf= pd.concat(frame, ignore_index=True)
                                    
                                    finaldf["slaoptions"]= pd.Categorical(finaldf["slaoptions"],categories=urutanstatus+["Total"], ordered=True)
                                    finaldf=finaldf.sort_values(["VendorName", "slaoptions"])
                                    finaldf= pd.concat(frame, ignore_index=True)
                                    listvendor= sorted(finaldf["VendorName"].unique().tolist())
                                    filtervendor= st.multiselect("Select Vendor", options= listvendor, default= listvendor, key="vendor_filter")
                                    if filtervendor:
                                        finaldf= finaldf[finaldf["VendorName"].isin(filtervendor)]
                                        
                                    grand_total=finaldf.loc[finaldf["slaoptions"]=="Total", urutanstatus+["TOTAL2","Total"]].sum().to_dict()
                                    grand_total["VendorName"]="Grand Total"
                                    grand_total["slaoptions"]= ""
                                    finaldf= pd.concat([finaldf, pd.DataFrame([grand_total])], ignore_index= True)
                                    
                            
                                    
                                    finaldf["VendorName"]= finaldf["VendorName"].mask(finaldf["VendorName"].duplicated(),"")
                                    barisheader=pd.MultiIndex.from_tuples([("", "Vendor"),("", "SLA"),("ONGOING-NOW", "OPEN"),("ONGOING-NOW", "ONPROGRESS"),("ONGOING-NOW", "POSTPONE"),("ONGOING-NOW", "TOTAL"),("OPEN-COMPLETE", "COMPLETE"),("", "INTEGRATION FAILED"),("COMP NOTE REQ & POSTPONE REQ", "APPROVAL DISPATCHER FS"), ("", "CANCEL"), ("", "TOTAL")], names=[None, None])
                                    finaldf=finaldf.reindex(columns=["VendorName", "slaoptions"]+urutanstatus[:3]+["TOTAL2"]+urutanstatus[3:]+["Total"])
                                    finaldf.columns=barisheader
                                    
                                    
              

                else:
                    st.warning("data kolomny galengkap di sheet historyworkorder")
                    
            else:
                st.warning ("ganemu historywo")
                
            dataframesla= tergabung_valid.copy()
            if uploaded is not None:
                out, filename= exportfile(uploaded)
            st.caption("Click the button below to calculate the status duration, status SLA and export as Excel")  
            # exportbutton, fillerexpbutton1, fillerexpbutton2= st.columns([1, 2, 3])
            # with exportbutton:
            if st.button("Export to Excel", type="primary", use_container_width= True):

                with st.spinner("Processing..."):
                    
                    if "finalcopy" in st.session_state:
                        final2= st.session_state.finalcopy.copy()
                        if{"WorkOrderNumber", "StatusReport"}.issubset(dataframesla.columns):
                            temp_df= dataframesla[["WorkOrderNumber", "StatusReport", "duration"]].copy()
                            temp_df= temp_df.rename(columns={"WorkOrderNumber":"WO Fieldsa"})
                            final2= final2.merge(temp_df,  on=["WO Fieldsa", "StatusReport"], how="left")
                            
                            statusreportmap={
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
                                "Posted To Ax Integration Success": "COMPLETE",
                                
                                "Postpone": "POSTPONE",
                                
                                "Sms Integration Failed": "INTEGRATION FAILED",
                                "Posted To Ax Integration Failed": "INTEGRATION FAILED",
                                "Provisioning Failed": "INTEGRATION FAILED",
                                
                                "Complete With Note Request": "APPROVAL DISPATCHER FS",
                                "Postpone Request": "APPROVAL DISPATCHER FS",
                                
                                "Cancel Work Order": "CANCEL"}
                            
                            def slaoptions_general(hour):
                                if pd.isna(hour):
                                    return None
                                if hour <= 4:
                                    return "0-4 Jam"
                                elif hour <= 6:
                                    return "4-6 Jam"
                                elif hour <= 12:
                                    return "6-12 Jam"
                                else:
                                    return ">12 Jam"

                            def slaoptions_broadband(hour):
                                if pd.isna(hour):
                                    return None
                                if hour <= 6:
                                    return "0-6 Jam"
                                elif hour <= 12:
                                    return "6-12 Jam"
                                elif hour <= 24:
                                    return "12-24 Jam"
                                else:
                                    return ">24 Jam"
                                
                            def get_sla(row):
                                div= str(row["DivisionName"])
                                dur_sla= row["duration"]
                                if "Broadband" in div:
                                    return slaoptions_broadband(dur_sla)
                                elif "LMS" in div or "Fiberisasi" in div:
                                    return slaoptions_general(dur_sla)
                                else:
                                    return None
                            final2["SLA Summary"]= final2.apply(get_sla, axis=1)
                            final2 = final2.drop(columns=["duration"])
                            # st.write("liat")
                            # st.dataframe(final2)
                            
                            buffer= BytesIO()
                            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                                final2.to_excel(writer, index=False, sheet_name= "SLA")
                                df_status= pd.DataFrame(list(statusreportmap.items()), columns=["Status", "Klasifikasi Status"])
                                df_status.to_excel(writer, index=False, sheet_name= "KeteranganStatus")
                            buffer.seek(0)
                            nameformat = f"Dashboard_FIELDSA_{pd.Timestamp.now():%Y%m%d_%H%M%S}.xlsx"
                            
                            st.success("Export ready")
                            st.download_button("Download Excel", data= buffer.getvalue(), file_name=nameformat, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True )
                        else:
                            st.write("gabisa")
                
                                
                                
                                
                                
                            
#     if menu_sidebar == "Data Comparation":
#             st.divider()
#             st.write("## Data Comparation")
#             st.write("Masih kosong :)")
    
    
    


    
# if menu_sidebar == "Merge Files":
#             st.divider()
#             st.write("## Merge Fieldsa Excel Files")
#             files_uploader= st.file_uploader("Upload Excel File to merge (min. 2)", type=["xlsx"], accept_multiple_files=True)



          
                            
         
                        
                        
                        
                        
                        
                        
                        
                        
    
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
