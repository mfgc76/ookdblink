/**************************************************************
 * 
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements.  See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership.  The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License.  You may obtain a copy of the License at
 * 
 *   http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied.  See the License for the
 * specific language governing permissions and limitations
 * under the License.
 * 
 *************************************************************/
/*
 * OfficeUNOClientApp1.java
 *
 * Created on 2016.07.10 - 07:15:58
 *
 */

package com.example;

import com.sun.star.beans.PropertyValue;
import com.sun.star.uno.XComponentContext;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.sheet.XSpreadsheetDocument; 
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.lang.XComponent; 
import com.sun.star.sheet.XSpreadsheets;
import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;
import java.io.IOException;
import java.lang.reflect.Array;
import java.net.URISyntaxException;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.TimeZone;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.swing.table.AbstractTableModel;

import kx.c;
import kx.c.KException;
import kx.c.Second;
/**
 *
 * @author mfitsilis
 */
public class OfficeUNOClientApp1 {
    
    /** Creates a new instance of OfficeUNOClientApp1 */
    public OfficeUNOClientApp1() {
    }
    //http://stackoverflow.com/questions/29684874/converting-a1-to-r1c1-format
    public static int r1c1x(String adr){ //get number equiv. to first letters
        int i = 0;
        int ret = 0;    
        while(adr.charAt(i) >= 'A' && adr.charAt(i) <= 'Z') {
            ret = i * 26 + (adr.charAt(i) - 'A');
            i++;
        }
        return ret;
    }
    public static int r1c1y(String adr){ //get 1st number in string
        int ret=0;
        Matcher match = Pattern.compile("[0-9]+").matcher(adr);
        if(match.find()) {
            ret=Integer.parseInt(match.group())-1; //row,col numbering start at 0
        }
        return ret;
    }
    //http://code.kx.com/wiki/Cookbook/InterfacingWithJava#Example_Grid_Viewer_using_Swing
    public static class KxTableModel extends AbstractTableModel {
        private c.Flip flip;
        private c.Dict dict;
        public void setDict(c.Dict data) {
            this.dict = data;
        }
        public void setFlip(c.Flip data) {
            this.flip = data;
        }
        public int getDRowCount() {
            return Array.getLength(dict.x);
        }
        public int getDColumnCount() {
            return Array.getLength(dict.y);
        }
        public int getRowCount() {
            return Array.getLength(flip.y[0]);
        }
        public int getColumnCount() {
            return flip.y.length;
        }
        public Object getValueAt(int rowIndex, int columnIndex) {
            return c.at(flip.y[columnIndex], rowIndex);
        }
        public String getColumnName(int columnIndex) {
            return flip.x[columnIndex];
        }
    };

    public static void closeatm(c c,XSpreadsheet xSpreadsheet,int destx,int desty,Object[][] tmp1) throws IOException, IndexOutOfBoundsException{
                    c.close();
                    XCellRange xrng;
                    xrng=xSpreadsheet.getCellRangeByPosition(destx,desty,destx,desty);
                    com.sun.star.sheet.XCellRangeData xData =(com.sun.star.sheet.XCellRangeData) UnoRuntime.queryInterface(com.sun.star.sheet.XCellRangeData.class, xrng);
                    xData.setDataArray(tmp1);
    }         
    public static void closelst(c c,XSpreadsheet xSpreadsheet,int destx,int desty,int len,int fp,Object[][] tmp1) throws IOException, IndexOutOfBoundsException{
                   c.close();
                   XCellRange xrng;
                   if(fp==0)
                        xrng=xSpreadsheet.getCellRangeByPosition(destx,desty,destx,desty+len-1);
                   else
                        xrng=xSpreadsheet.getCellRangeByPosition(destx,desty,destx+len-1,desty);
                   com.sun.star.sheet.XCellRangeData xData =(com.sun.star.sheet.XCellRangeData) UnoRuntime.queryInterface(com.sun.star.sheet.XCellRangeData.class, xrng);
                   xData.setDataArray(tmp1);
    }         
    public static void qexec(String host, String query, String destination, int header, int flip, int refresh, XSpreadsheet xSpreadsheet){   
        KxTableModel model = new KxTableModel();
        c c = null;
        String lhost;
        String hoststr[];
        hoststr=host.split(":");  //split at ":"  hostname:port:(username:password), user,passwd are optional
        int hd=header; //show table header
        int fp=flip; //rotate list/table
        int nrow,ncol,len;
        String user,pass;
        XCellRange xrng;
        int destx=r1c1x(destination.split(":")[1]); //destination position x for the result
        int desty=r1c1y(destination.split(":")[1]);
                
        if(hoststr[0].equals("")) lhost="localhost";
        else lhost=hoststr[0];
        int lport=Integer.parseInt(hoststr[1]);

        if(!query.equals("")){
            try {
                    if(hoststr.length==4){
                        user=hoststr[2];
                        pass=hoststr[3];
                        c = new c(lhost,lport,(user+":"+pass));
                    }
                    else
                        c = new c(lhost,lport);
                    
                    c.tz=TimeZone.getTimeZone("GMT"); //set timezone to gmt
            } catch (java.lang.Exception ex) {
                    System.err.println (ex);
            }
         Object res;
         String[] strs;
         Date[] dats;
         long[] lngs;
         Second[] scnds;
         double[] dbls;
         //http://stackoverflow.com/questions/21680618/how-to-iterate-through-an-object-that-is-pointing-to-array-of-doubles#21680650         
         Object [][] tmp1=new Object[1][1]; //row x col
         try {
                res= c.k(query);
                //atoms
                if(res instanceof String){
                    tmp1[0][0]=(String)res;
                    closeatm(c,xSpreadsheet,destx,desty,tmp1);
                }
                else if(res instanceof Date){
                    tmp1[0][0]=(Double)((Long) ((Date)(res)).getTime() ).doubleValue();
                    closeatm(c,xSpreadsheet,destx,desty,tmp1);
                }
                else if(res instanceof Long){
                    tmp1[0][0]=(Double)((Long) (res) ).doubleValue();
                    closeatm(c,xSpreadsheet,destx,desty,tmp1);
                }
                else if(res instanceof Second){
                    tmp1[0][0]=(Double)(double)((Second)(res)).i;
                    closeatm(c,xSpreadsheet,destx,desty,tmp1);
                }
                else if(res instanceof Double){
                    tmp1[0][0]=res;
                    closeatm(c,xSpreadsheet,destx,desty,tmp1);
                }
                //lists
                else if(res instanceof String[]){
                  strs=(String[])res;
                  nrow=len=strs.length;ncol=1;
                  if (fp==1) {int tmp=nrow; nrow=ncol; ncol=tmp; } 
                  Object [][] tmp2=new Object[nrow][ncol]; //row x col
                  for(int i=0;i<strs.length;i++){
                      if (fp==0)tmp2[i][0]=strs[i]; 
                      else tmp2[0][i]=strs[i];
                  }
                  closelst(c,xSpreadsheet,destx,desty,len,fp,tmp2);
                }
                else if(res instanceof Date[]){
                  dats=(Date[])res;
                  nrow=len=dats.length;ncol=1;
                  if (fp==1) {int tmp=nrow; nrow=ncol; ncol=tmp; } 
                  Object [][] tmp2=new Object[nrow][ncol]; //row x col
                  for(int i=0;i<dats.length;i++){
                      if (fp==0)tmp2[i][0]=(Double)((Long) (dats[i]).getTime() ).doubleValue();
                      else tmp2[0][i]=(Double)((Long) (dats[i]).getTime() ).doubleValue();
                  }
                  closelst(c,xSpreadsheet,destx,desty,len,fp,tmp2);
                }
                else if(res instanceof long[]){
                  lngs=(long[])res;
                  nrow=len=lngs.length;ncol=1;
                  if (fp==1) {int tmp=nrow; nrow=ncol; ncol=tmp; } 
                  Object [][] tmp2=new Object[nrow][ncol]; //row x col
                  for(int i=0;i<lngs.length;i++){
                      if (fp==0)tmp2[i][0]=(Double)(Long.valueOf(i)).doubleValue(); 
                      else tmp2[0][i]=(Double)(Long.valueOf(i)).doubleValue(); 
                  }
                  closelst(c,xSpreadsheet,destx,desty,len,fp,tmp2);
                }
                else if(res instanceof Second[]){
                  scnds=(Second[])res;
                  nrow=len=scnds.length;ncol=1;
                  if (fp==1) {int tmp=nrow; nrow=ncol; ncol=tmp; } 
                  Object [][] tmp2=new Object[nrow][ncol]; //row x col
                  for(int i=0;i<scnds.length;i++){
                      if (fp==0)tmp2[i][0]=(Double)(double)(scnds[i]).i;
                      else tmp2[0][i]=(Double)(double)(scnds[i]).i;
                  }
                  closelst(c,xSpreadsheet,destx,desty,len,fp,tmp2);
                }
                else if(res instanceof double[]){
                  dbls=(double[])res;
                  nrow=len=dbls.length;ncol=1;
                  if (fp==1) {int tmp=nrow; nrow=ncol; ncol=tmp; } 
                  Object [][] tmp2=new Object[nrow][ncol]; //row x col
                  for(int i=0;i<dbls.length;i++){
                      if (fp==0)tmp2[i][0]=(Double)dbls[i];
                      else tmp2[0][i]=(Double)dbls[i];
                  }
                  closelst(c,xSpreadsheet,destx,desty,len,fp,tmp2);
                }
                else if(res instanceof c.Dict){ //todo complete...
                        model.setDict((c.Dict) c.k(query));
                        nrow=model.getDRowCount(); //key
                        ncol=model.getDColumnCount(); //value ([] or object[])
                        Object [][] tmp4=new Object[nrow][ncol];
                      /*  for(int i=0;i<nrow-hd;i++){
                        for(int j=0;j<ncol;j++){
                            
                        }
                        }
                       */
                }
                //tables
                else if(res instanceof c.Flip){
                    try {
                        model.setFlip((c.Flip) c.k(query));
                        nrow=model.getRowCount();
                        ncol=model.getColumnCount();

                        //must use 0! if table is keyed
                        if (fp==1) {int tmp=nrow; nrow=ncol; ncol=tmp; } // cannor swap without tmp variable 
                        if(hd==1) {if (fp==0) nrow++; else ncol++; } //include header
                        Object [][] tmp3=new Object[nrow][ncol]; //row x col
                        if(hd==1){
                            for(int i=0;i<(fp==0?ncol:nrow);i++){  //first the header
                            //tmp3[0][i]= model.getColumnName(i);
                                if (fp==0)tmp3[0][i]=model.getColumnName(i);
                                else tmp3[i][0]=     model.getColumnName(i);
                            }
                        }                       
                        if (fp==1) {int tmp=nrow; nrow=ncol; ncol=tmp; } //swap again for assignment
                        for(int i=0;i<nrow-hd;i++){
                        for(int j=0;j<ncol;j++){
                        if(model.getValueAt(i,j) instanceof String){
                        //tmp3[i+hd][j]=model.getValueAt(i,j);
                            if (fp==0)tmp3[i+hd][j]=model.getValueAt(i,j);
                            else tmp3[j][i+hd]=     model.getValueAt(i,j);
                        }
                        else if(model.getValueAt(i,j) instanceof Date){
                        //tmp3[i+hd][j]=(Double)((Long) ((Date)(model.getValueAt(i,j))).getTime() ).doubleValue();
                            if (fp==0)tmp3[i+hd][j]=(Double)((Long) ((Date)(model.getValueAt(i,j))).getTime() ).doubleValue();
                            else tmp3[j][i+hd]=     (Double)((Long) ((Date)(model.getValueAt(i,j))).getTime() ).doubleValue();
                        }
                        else if(model.getValueAt(i,j) instanceof Long){
                        //tmp3[i+hd][j]=(Double)((Long)(model.getValueAt(i,j))).doubleValue();
                            if (fp==0)tmp3[i+hd][j]=(Double)((Long)(model.getValueAt(i,j))).doubleValue();
                            else tmp3[j][i+hd]=     (Double)((Long)(model.getValueAt(i,j))).doubleValue();
                        }
                        else if(model.getValueAt(i,j) instanceof Second){
                        //tmp3[i+hd][j]= (double)((Second)(model.getValueAt(i,j))).i ;
                            if (fp==0)tmp3[i+hd][j]=(double)((Second)(model.getValueAt(i,j))).i;
                            else tmp3[j][i+hd]=     (double)((Second)(model.getValueAt(i,j))).i;
                        }
                        else if(model.getValueAt(i,j) instanceof Double){
                        //tmp3[i+hd][j]=model.getValueAt(i,j);
                            if (fp==0)tmp3[i+hd][j]=model.getValueAt(i,j);
                            else tmp3[j][i+hd]=     model.getValueAt(i,j);
                        }
                        else{
                        //if(qn(model.getValueAt(i,j))==true)
                        //tmp3[i+hd][j]="";   //works with null
                            if (fp==0)tmp3[i+hd][j]="";
                            else tmp3[j][i+hd]=     "";
                        }
                        }
                        }
                        c.close();
                        if(fp==0)
                            xrng=xSpreadsheet.getCellRangeByPosition(destx,desty,destx+ncol-1,desty+nrow-1);
                        else
                            xrng=xSpreadsheet.getCellRangeByPosition(destx,desty,destx+nrow-1,desty+ncol-1);
                        com.sun.star.sheet.XCellRangeData xData =(com.sun.star.sheet.XCellRangeData) UnoRuntime.queryInterface(com.sun.star.sheet.XCellRangeData.class, xrng);
                        xData.setDataArray(tmp3); //set Object[][]
                       
                        } catch (KException e1) {
                        } catch (IOException ex) {
                        } catch (IndexOutOfBoundsException ex) {
                        Logger.getLogger(OfficeUNOClientApp1.class.getName()).log(Level.SEVERE, null, ex);
                    }
                    
                }

            } catch (KException ex) {
            } catch (IOException ex) {
            } catch (IndexOutOfBoundsException ex) {
                Logger.getLogger(OfficeUNOClientApp1.class.getName()).log(Level.SEVERE, null, ex);
            }
        }       
        //return new Object[0][0];
   }
    /**
     * @param args the command line arguments
     */ 
    public static void main(String[] args) throws URISyntaxException {//todo fix...
        //https://wiki.openoffice.org/wiki/Documentation/DevGuide/OfficeDev/Using_the_Desktop
        //from https://wiki.openoffice.org/wiki/Documentation/DevGuide/FirstSteps/Example:_Working_with_a_Spreadsheet_Document

    //    System.out.println("args:"+Arrays.toString(args));
        String[] dashstrname=Arrays.toString(args).replace("[", "").replace("]", "").split(",");  //get 1st argument
        String setstr="settings"; //default name of settings tab
        String entrstr="J1"; //default cell position for num of entries
        if (dashstrname.length==3){
            setstr=dashstrname[1];
            entrstr=dashstrname[2];
        }
        if(dashstrname[0].equals("")) dashstrname[0]="dashloader";
        
    //    System.out.println(dashstrname[0]);System.exit(0);        
        try {
            //dashstrname[0]="test";
            //dashstrname[0]="dashloader";
            
            //System.out.println(dashstrname[0]);    
            //if(!dashstrname[0].equals("]"))
            {
            // get the remote office component context
            XComponentContext xRemoteContext = Bootstrap.bootstrap();
            if (xRemoteContext == null) {
                System.err.println("ERROR: Could not bootstrap default Office.");
            }
 
            XMultiComponentFactory xRemoteServiceManager = xRemoteContext.getServiceManager();
 
            Object desktop = xRemoteServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", xRemoteContext);
            XComponentLoader xComponentLoader = (XComponentLoader)  UnoRuntime.queryInterface(XComponentLoader.class, desktop);
          //  XComponentLoader xComponentLoader2 = (XComponentLoader) UnoRuntime.queryInterface(XComponentLoader.class, desktop);
 
            PropertyValue[] loadProps = new PropertyValue[0];
            XComponent xSpreadsheetComponent;
            
            //if(dashstrname[0].equals(""))
            //    xSpreadsheetComponent= xComponentLoader.loadComponentFromURL("private:factory/scalc", "_blank", 0, loadProps);
            //else
            String path; //http://stackoverflow.com/questions/320542/how-to-get-the-path-of-a-running-jar-file
            path=OfficeUNOClientApp1.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath();
            path="file://"+path.replace("\\", "/")+"/../"+dashstrname[0]+".ods";
        //    System.out.println(path);
            xSpreadsheetComponent = xComponentLoader.loadComponentFromURL(path, "_blank", 0, loadProps);
            
//                xSpreadsheetComponent = xComponentLoader.loadComponentFromURL("file:///C:/Users/mfitsilis/Documents/myspreadsheets/"+dashstrname[0]+".ods", "_blank", 0, loadProps);
            //if filename not exist, exits at finally{}
            XSpreadsheetDocument xSpreadsheetDocument = (XSpreadsheetDocument) UnoRuntime.queryInterface(XSpreadsheetDocument.class,xSpreadsheetComponent);
            XSpreadsheets xSpreadsheets = xSpreadsheetDocument.getSheets();
            String[] sprdnames=xSpreadsheets.getElementNames();
            XCell xCell,xCell2;
            int totnumofq; //total number of queries starting from row 2(row 1 is header)
            int reload=1; //reload flag is reset to 1 after reloading
            long curtimer; //get current time in seconds since epoch
            //calc timezone offset to display localtime in oocalc http://stackoverflow.com/questions/11399491/java-timezone-offset
            TimeZone tz = TimeZone.getTimeZone("UTC");
            tz = TimeZone.getDefault();  
            Calendar cal=Calendar.getInstance(TimeZone.getTimeZone("tz"));
            int offsetInMillis = tz.getOffset(cal.getTimeInMillis());        
               
            //if connection times out it might be the wrong host:port
            //http://stackoverflow.com/questions/5662283/java-net-connectexception-connection-timed-out-connect  
            if(Arrays.asList(sprdnames).contains(setstr))
            //host:port:user:pass,query,destination,header,flip,refresh(s)
            {
                Object sheet = xSpreadsheets.getByName(setstr);
                XSpreadsheet xSpreadsheet = (XSpreadsheet)UnoRuntime.queryInterface(XSpreadsheet.class, sheet);
                totnumofq=(int)(xSpreadsheet.getCellByPosition(r1c1x(entrstr),r1c1y(entrstr))).getValue();
                (xSpreadsheet.getCellByPosition(r1c1x(entrstr),r1c1y(entrstr)+1)).setValue(0); //reset reload to 0
                String[] host=new String[(int) totnumofq];
                String[] query=new String[(int) totnumofq];
                String[] destination=new String[(int) totnumofq];
                Integer[] header=new Integer[(int) totnumofq];
                Integer[] flip=new Integer[(int) totnumofq];
                Integer[] refresh=new Integer[(int) totnumofq];
                Long[] timer=new Long[(int) totnumofq];
                curtimer=(System.currentTimeMillis()+offsetInMillis) / 1000l;                 
                
                for(int i=0;i<totnumofq;i++){
                    host[i]=        (xSpreadsheet.getCellByPosition(0,i+1)).getFormula();
                    query[i]=       (xSpreadsheet.getCellByPosition(1,i+1)).getFormula();
                    destination[i]= (xSpreadsheet.getCellByPosition(2,i+1)).getFormula();
                    header[i]=(int) (xSpreadsheet.getCellByPosition(3,i+1)).getValue();
                    flip[i]=(int)   (xSpreadsheet.getCellByPosition(4,i+1)).getValue();
                    refresh[i]=(int)(xSpreadsheet.getCellByPosition(5,i+1)).getValue();
                    timer[i]=curtimer+refresh[i];
                    (xSpreadsheet.getCellByPosition(r1c1x(entrstr),r1c1y(entrstr)+2)).setValue(curtimer); //set update time for settings sheet
                }
            for(;;) //endless loop
            {
                //XSpreadsheets xSpreadsheets = xSpreadsheetDocument.getSheets();
                //String[] sprdnames=xSpreadsheets.getElementNames();
                //if(sprdnames.equals("settings")){
                //for(int i=0;i<sprdnames.length;i++){
                reload=(int)(xSpreadsheet.getCellByPosition(r1c1x(entrstr),r1c1y(entrstr)+1)).getValue();
                if(reload==1){
                    totnumofq=(int)(xSpreadsheet.getCellByPosition(r1c1x(entrstr),r1c1y(entrstr))).getValue();
                    (xSpreadsheet.getCellByPosition(r1c1x(entrstr),r1c1y(entrstr)+1)).setValue(0); //reset to 0
                    host=new String[(int) totnumofq];
                    query=new String[(int) totnumofq];
                    destination=new String[(int) totnumofq];
                    header=new Integer[(int) totnumofq];
                    flip=new Integer[(int) totnumofq];
                    refresh=new Integer[(int) totnumofq];
                    timer=new Long[(int) totnumofq];
                    curtimer=(System.currentTimeMillis()+offsetInMillis) / 1000l;                 
                    (xSpreadsheet.getCellByPosition(r1c1x(entrstr),r1c1y(entrstr)+2)).setValue(curtimer); //set update time for settings sheet
                       for(int i=0;i<totnumofq;i++){
                           host[i]=        (xSpreadsheet.getCellByPosition(0,i+1)).getFormula();
                           query[i]=       (xSpreadsheet.getCellByPosition(1,i+1)).getFormula();
                           destination[i]= (xSpreadsheet.getCellByPosition(2,i+1)).getFormula();
                           header[i]=(int) (xSpreadsheet.getCellByPosition(3,i+1)).getValue();
                           flip[i]=(int)   (xSpreadsheet.getCellByPosition(4,i+1)).getValue();
                           refresh[i]=(int)(xSpreadsheet.getCellByPosition(5,i+1)).getValue();
                           timer[i]=curtimer+refresh[i];
                       }
                }
                for(int i=0;i<totnumofq;i++){ //do for every query  
                    curtimer=(System.currentTimeMillis()+offsetInMillis) / 1000l;                 
                    if(curtimer>timer[i]){
                        Object sheet2 = xSpreadsheets.getByName(destination[i].split(":")[0]);
                        XSpreadsheet xSpreadsheet2 = (XSpreadsheet)UnoRuntime.queryInterface(XSpreadsheet.class, sheet2);

                        qexec(host[i],query[i],destination[i],(int)header[i],(int)flip[i],(int)refresh[i],xSpreadsheet2);
                        timer[i]+=refresh[i];  //set next timer expiry
                        (xSpreadsheet.getCellByPosition(6,i+1)).setValue(curtimer); //set update time for query

                    }
                }
                Thread.sleep(200);
            } //end for(;;)
            } //if tab "settings" exist else exit program
        }//end of dashstrname=="]"
        }//end of try
        catch (BootstrapException e){
        } catch (Exception e) {
        } catch (InterruptedException e) {
        }
        finally {
            System.exit( 0 );
        }   
    }
}
