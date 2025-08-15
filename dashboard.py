import streamlit as st 
import matplotlib.pyplot as plt
import plotly.express as px
import pandas as pd
from io import BytesIO
import numpy as np
import time
import os


st.set_page_config(layout="wide", page_title="FIELDSA DASHBOARD")
with open("header.html", "r") as head:
    st.markdown(head.read(), unsafe_allow_html=True)
    
    
uploaded=st.file_uploader("Upload Excel File", type=["xlsx", "csv"])

if uploaded is not None:
    @st.cache_data
    def load(file):
        if file.name.endswith(".csv"):
            return {"WorkOrder": pd.read_csv(file)}
        else:
            return pd.read_excel(file, sheet_name=None)
    exceldata=load(uploaded)
    df= exceldata.get("WorkOrder")
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    st.write("## Status Duration & SLA Calculation")
    st.caption("Click the button below to calculate the status duration, status SLA and export as Excel")  
     
    # def namafileauto(base, ext=".xlsx"):
    # i=1
    # filename=f"{base}({i}){ext}"
    # while os.path.exists(filename):
    #     i+=1
    #     filename=f"{base}({i}){ext}"
    # return filename

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



        urutanyangkumau= [
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

        final= final[[col for col in urutanyangkumau if col in final.columns]]

        buffer= BytesIO()
        with pd.ExcelWriter(buffer, engine= "xlsxwriter") as writer:
            final.to_excel(writer, index= False, sheet_name= "SLA")
        buffer.seek(0)
        suggestname = f"AllTaskList_FIELDSA_{pd.Timestamp.now():%Y%m%d_%H%M%S}.xlsx"
        return buffer, suggestname
    
    pikachu, charizard= st.columns([1, 2])
    with pikachu:
        if st.button("Export to Excel", type="primary", use_container_width= True):
            with st.spinner("Processing..."):
                out, filename= exportfile(uploaded)
            if out is not None:
                st.success("Export ready")
                st.download_button("Download Excel", data= out.getvalue(), file_name= filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True )
            else:
                st.warning("Export Failed")














    
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
        col1,col2=st.columns(2)
        col1.metric("Total wo", len(df))
        try:
            identifydate= uploaded.name.split("_")[1][:8]
            filedate= pd.to_datetime(identifydate, format="%Y%m%d")
            lastupdate= filedate.strftime("%d %B %Y")
        except:
            lastupdate= "Date not valid"
        
        # Last Update via max date in date range
        # lastupdate= df["Created"].max().strftime("%d %B %Y")
        
        col2.metric("Last Data Update", lastupdate)
        
        dfstat= df.copy() 
        
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
            loc= st.selectbox("Location", options=location, format_func= lambda x: "Sub Region" if x== "SubRegion" else x, key= "locfilter")
            loccol= "City" if "City" in loc else loc
            if loc=="City [Top 10]":
                top10= df["City"].value_counts().nlargest(10).index.tolist()
                df= df[df["City"].isin(top10)]
                
        with filter2:
            st.markdown("Division")
            with st.expander("Division"):
                divwo= ["Broadband","Lms","Fiberisasi"]
                jumlahperdiv= df["DivisionName"].value_counts().to_dict()
            
                divisiterpilih=[]
            
                for c in divwo:
                    itung= jumlahperdiv.get(c,0)
                    if st.checkbox(f"{c} ({itung})", value= (c in st.session_state.divfilter), key=f"{c}"):
                        divisiterpilih.append(c)
                st.session_state.divfilter= divisiterpilih
                
        with filter3:
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
    
   
        if divisiterpilih:
            df= df[df["DivisionName"].isin(divisiterpilih)]
        else:
            st.warning("No division type chosen")
            
        if tipeygdipilih:
            df= df[df["WorkOrderTypeName"].isin(tipeygdipilih)]
        else:
            st.warning("No Workorder type chosen")
   
        
        
        if loc:
            normcol= df[loccol]
            itungisi= normcol.value_counts().reset_index()
            itungisi.columns=[loccol, "Count"]
            
            bar= px.bar(
                 itungisi,
                x= loccol,
                y= "Count",
                color= loccol,
                title= f"Based on {ubahheader.get(loccol, loccol)}",
                labels= {loccol: ubahheader.get(loccol,loccol), "Count": "Jumlah"}
            )
                
            st.plotly_chart(bar, use_container_width=True)
        else:
            st.warning ("No column available for visualization")

        tambahcols=[loccol,"CustomerName", "VendorName", "Reason"]
        colkotak=st.columns(len(tambahcols))
        
        for index, namadicol in enumerate(tambahcols):
            displayindex= index+1
            with colkotak[index]:
                if namadicol in df.columns:
                    addcoldata= df[namadicol].value_counts().reset_index()
                    addcoldata.columns= [namadicol, "Amount"]
                    addcoldata.index= range(1, len(addcoldata) +1)
                    st.markdown(f"**{namadicol}**")
                    st.dataframe(addcoldata)
                else:
                    st.warning("Kolom tidak ditemukan")
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
                
   
        st.write("## Work Order Status Chart")   
        
        dfstat["WorkOrderStatusItem"] = dfstat["WorkOrderStatusItem"].astype(str).str.strip().str.title()

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
            "Posted To AX Integration Failed": "INTEGRATION FAILED",
            "Provisioning Failed": "INTEGRATION FAILED",
            "Cancel Work Order": "CANCEL",
            "Posted To AX Integration Success": "COMPLETE" }
        
        dfstat["StatusReport"]= dfstat["WorkOrderStatusItem"].map(statusmap).fillna("N/A")
        
        locfilter2,statfilter=st.columns(2)
        
        if "location2" not in st.session_state:
            st.session_state.location2= "Region"
        if "status" not in st.session_state:
            st.session_state.status= ["OPEN", "COMPLETE", "ONPROGRESS", "POSTPONE", "INTEGRATION FAILED", "CANCEL"]
 
        with locfilter2:
            loc= st.selectbox("Location", options=location, format_func= lambda x: "Sub Region" if x== "SubRegion" else x, key= "location2")
            loccol= "City" if "City" in loc else loc
            if loc=="City [Top 10]":
                top10= dfstat["City"].value_counts().nlargest(10).index.tolist()
                dfstat= dfstat[dfstat["City"].isin(top10)]

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
            st.warning("No status chosen")
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
            statussummary["GRAND TOTAL"]= statussummary.loc[:, statussummary.columns != loccol].sum(axis=1)
            totalperstatus= statussummary.loc[:, statussummary.columns != loccol].sum(axis=0)
            totalbaris=pd.DataFrame([totalperstatus])
            totalbaris[loccol]= "GRAND TOTAL"
            statussummary= pd.concat([statussummary, totalbaris], ignore_index= True)
            statussummary.index= range(1, len(statussummary)+1)
            st.dataframe(statussummary)
        else:
            st.warning("No data available")
        
#Nyobacommit
      
