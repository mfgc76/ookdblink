# ookdblink
openoffice link to kdb to create a dashboard

<p>This is a stand alone app than takes the name of a
spreadsheet as a parameter and uses the settings sheet inside
to make queries to kdb and update cells(polling) in selected
cells and sheets.</p>

The jar executable takes 2 parameters: the spreasheet filename without extension and the cellname of the reload cell (eg. J1) 

- username,password are optional,empty host is localhost
- header: 0:no header, 1:show header if result is table
- flip: 0:no rotation, 1:rotate if result is table or list
- switch reload cell to 1 to refresh the queries list
  it is periodically refreshed and will switch back to 0
  unlike ookdbaddin where the formula does not allow an update
  or a longer table, the tables update without such issues 
- set cell over reload cell to monitor the number of query entries eg.: =COUNTIF(A1:A10000;"<>""")-1
- supported time formats: time, date, second return milliseconds
  since epoch (UTC), better use formula (G2 / 86400000) + DATE(1970;1;1) to manually convert, as in the example below, to oocalc format and format the cell accordingly, unsupported formats return empty cell
- in the example below the header in the 1st row shows the column order

<p>Example</p>
![<oocalc image>](https://github.com/mfitsilis/ookdblink/blob/master/img/ookdblink1.png)

- To run use runjar.bat and set the correct path to file ookdblink.jar inside dist directory and arguments
- To build import project into Netbeans, install the openoffice [plug-in](https://wiki.openoffice.org/wiki/OpenOffice_NetBeans_Integration#NetBeans_8.x_and_Apache_OpenOffice_4.1.x) first
- Built with Java7, tested with kdb+ 3.2, Openoffice 4.11
