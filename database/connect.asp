<%
    dim connectstring
    connectstring = "driver={MariaDB ODBC 3.1 Driver};server=www.skyflight.cn;database=OMS;uid=root;pwd=549268Mar"
%>
<%
    function connectstr(driver,serveraddr,database,uid,pwd)
        dim fcnn
        fcnn="driver={" & driver &"};server=" & serveraddr & ";database=" & database &";uid=" & uid & ";pwd=" & pwd
        connectstr = fcnn
    end function
%>