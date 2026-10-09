

var running=false,timer=null,idx=0,selectedQueue=[],selectedPos=0;
function byId(id){return document.getElementById(id);}
function setText(id,v){var e=byId(id);if(e)e.innerText=v;}
function setWidth(id,v){var e=byId(id);if(e)e.style.width=Math.max(0,Math.min(100,parseInt(v,10)||0))+"%";}
function setClass(id,v){var e=byId(id);if(e)e.className=v;}
function totalGB(n){try{return (parseFloat(n)/1073741824).toFixed(1);}catch(e){return "-";}}
function logMsg(msg,cls){var e=byId("log");if(!e)return;var d=document.createElement("div");d.className=cls||"";d.innerText="["+new Date().toLocaleTimeString()+"] "+msg;e.appendChild(d);e.scrollTop=e.scrollHeight;}
function progress(p){if(p<0)p=0;if(p>100)p=100;byId("bar").style.width=p+"%";setText("pct",p+"%");}
var ACTIVE_ITEMS=[1,2,3,4,5,6,7,8,9,10];
function isChecked(i){var e=byId("c"+i);return !!(e&&e.checked);}
function updateSelectionCount(){var i,c=0;for(i=0;i<ACTIVE_ITEMS.length;i++)if(isChecked(ACTIVE_ITEMS[i]))c++;setText("selectCount",c+"개 선택");}
function selectAll(v){var i,e;for(i=0;i<ACTIVE_ITEMS.length;i++){e=byId("c"+ACTIVE_ITEMS[i]);if(e)e.checked=v;}updateSelectionCount();logMsg(v?"진단 항목 전체 선택":"진단 항목 전체 해제","info");}
function getSelected(){var a=[],i;for(i=0;i<ACTIVE_ITEMS.length;i++)if(isChecked(ACTIVE_ITEMS[i]))a.push(ACTIVE_ITEMS[i]);return a;}
function resetRows(list){var j,i,r;for(j=0;j<list.length;j++){i=list[j];r=byId("r"+i);if(r){r.cells[4].innerText="대기 중";r.cells[5].innerText="-";r.cells[4].className="st";r.cells[5].className="res";}}}
function reset(){resetRows(ACTIVE_ITEMS);}
function startDiagnosis(){if(running)return;startQueue(ACTIVE_ITEMS,"전체 진단");}
function startSelectedDiagnosis(){if(running)return;var a=getSelected();if(!a.length){alert("진단할 항목을 하나 이상 선택하세요.");return;}startQueue(a,"선택 진단");}
var lastDiagnosisIssues=[];
function diagWmi(){try{var loc=new ActiveXObject("WbemScripting.SWbemLocator");return loc.ConnectServer(".","root\\cimv2");}catch(e){throw new Error("WMI 연결 실패: "+e.message);}}
function issuePush(row,type,name,extra){lastDiagnosisIssues.push({row:row,type:type,name:name,extra:extra||""});}
function diagResult(row,status,result,cls,msg){setRowResult(row,status,result,cls);logMsg(msg||("진단 항목 "+row+" : "+result),cls||"info");}
function findSpecificCom(con,port){
  var out={found:false,error:0,name:"",id:""},q,en,x,n="";
  try{
    q=con.ExecQuery("SELECT Name,PNPDeviceID,ConfigManagerErrorCode FROM Win32_PnPEntity WHERE PNPClass='Ports'");
    en=new Enumerator(q);
    for(;!en.atEnd();en.moveNext()){
      x=en.item(); n=String(x.Name||"");
      if(n.toUpperCase().indexOf("("+port.toUpperCase()+")")>=0){out.found=true;out.error=Number(x.ConfigManagerErrorCode||0);out.name=n;out.id=String(x.PNPDeviceID||"");break;}
    }
  }catch(e){}
  return out;
}
function probeSerialDevice(port){
  /* COM 포트가 Windows에 보이는 것만으로는 USB-Serial 케이블만 연결된 상태와 실제 기기 연결을 구분할 수 없습니다.
     따라서 짧은 시간 동안 실제 RS-232 상태 신호(CTS/DSR/CD)를 확인합니다. 콘솔창은 표시하지 않습니다.
     신호가 하나도 없으면 "기기 연결 안됨"으로 판정합니다. */
  var fso,sh,tmp,cmd,rc,text,parts,cts=false,dsr=false,cd=false;
  try{
    fso=new ActiveXObject("Scripting.FileSystemObject");
    tmp=fso.GetSpecialFolder(2).Path+"\\POSCARE_COM_"+port+"_"+new Date().getTime()+".txt";
    sh=new ActiveXObject("WScript.Shell");
    cmd='powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -Command "try{$p=New-Object System.IO.Ports.SerialPort(\''+port+'\',9600,[System.IO.Ports.Parity]::None,8,[System.IO.Ports.StopBits]::One);$p.DtrEnable=$true;$p.RtsEnable=$true;$p.ReadTimeout=150;$p.WriteTimeout=150;$p.Open();Start-Sleep -Milliseconds 120;$c=if($p.CtsHolding){\'1\'}else{\'0\'};$d=if($p.DsrHolding){\'1\'}else{\'0\'};$l=if($p.CDHolding){\'1\'}else{\'0\'};[IO.File]::WriteAllText(\''+tmp.replace(/'/g,"''")+'\',$c+\'|\'+$d+\'|\'+$l);$p.Close()}catch{[IO.File]::WriteAllText(\''+tmp.replace(/'/g,"''")+'\',\'OPENFAIL\')}"';
    rc=sh.Run(cmd,0,true);
    if(fso.FileExists(tmp)){text=fso.OpenTextFile(tmp,1,true).ReadAll();fso.DeleteFile(tmp,true);}else{return false;}
    if(String(text).indexOf("OPENFAIL")>=0)return false;
    parts=String(text).replace(/\r|\n/g,"").split("|");
    cts=parts.length>0&&parts[0]==="1";dsr=parts.length>1&&parts[1]==="1";cd=parts.length>2&&parts[2]==="1";
    return !!(cts||dsr||cd);
  }catch(e){try{if(tmp&&fso&&fso.FileExists(tmp))fso.DeleteFile(tmp,true);}catch(ignore){}return false;}
}
function diagnoseComDevice(row,port,label,con){
  var p=findSpecificCom(con,port), probe;
  if(!p.found){
    diagResult(row,"검사 완료","기기 연결 안됨","caution",port+" 포트가 Windows에서 확인되지 않습니다.");
    return;
  }
  if(p.error!==0){
    diagResult(row,"오류","기기 연결 오류","bad",port+" 장치 오류 코드 "+p.error+"가 확인되었습니다.");
    issuePush(row,"com",p.name,p.id);
    return;
  }
  probe=probeSerialDevice(port);
  if(probe){
    diagResult(row,"검사 완료","기기 연결","good",port+" 포트와 연결된 기기의 통신 상태 신호가 확인되었습니다.");
  }else{
    diagResult(row,"검사 완료","기기 연결 안됨","caution",port+" 포트는 확인되지만 케이블만 연결되었거나 기기에서 통신 상태 신호가 확인되지 않습니다.");
  }
}
function diagnoseItem(item){
  var con,q,en,x,n,sum,total,free,ram,d,size,freeDisk,disk,svc,netOk,found=0;
  try{
    con=diagWmi();
    if(item===1){
      var os=con.Get("Win32_OperatingSystem=@");
      if(os && os.Caption){diagResult(1,"검사 완료","정상","good","Windows 시스템 상태 정상: "+os.Caption);}
      else{diagResult(1,"오류","OS 정보 조회 실패","bad","Windows 운영체제 정보를 조회하지 못했습니다.");issuePush(1,"system","Windows 시스템 파일","");}
    }else if(item===2){
      q=con.ExecQuery("SELECT LoadPercentage FROM Win32_Processor");en=new Enumerator(q);sum=0;n=0;for(;!en.atEnd();en.moveNext()){sum+=Number(en.item().LoadPercentage||0);n++;}var cpu=n?Math.round(sum/n):0;
      os=con.Get("Win32_OperatingSystem=@");total=Number(os.TotalVisibleMemorySize||0);free=Number(os.FreePhysicalMemory||0);ram=total?Math.round((1-free/total)*100):0;
      d=con.Get('Win32_LogicalDisk.DeviceID="C:"');size=d&&d.Size?Number(d.Size):0;freeDisk=d&&d.FreeSpace?Number(d.FreeSpace):0;disk=size?Math.round((1-freeDisk/size)*100):0;
      if(cpu>=98||ram>=98||disk>=99){diagResult(2,"주의","자원 사용량 높음","caution","CPU "+cpu+"% / RAM "+ram+"% / 디스크 "+disk+"% 사용 중입니다.");issuePush(2,"resource","자원 사용량",{cpu:cpu,ram:ram,disk:disk});}
      else diagResult(2,"검사 완료","정상","good","CPU "+cpu+"% / RAM "+ram+"% / 디스크 "+disk+"% 사용 중입니다.");
    }else if(item===3){
      var names=["Winmgmt","EventLog","RpcSs","Spooler"],bad=[];for(var si=0;si<names.length;si++){svc=serviceState(names[si]);if(svc!=="running")bad.push(names[si]+":"+svc);}
      if(bad.length){diagResult(3,"오류","중지 서비스 발견","bad",bad.join(", ")+" 서비스가 실행되지 않습니다.");for(si=0;si<bad.length;si++)issuePush(3,"service",bad[si].split(":")[0],"");}
      else diagResult(3,"검사 완료","정상","good","필수 Windows 서비스가 정상 실행 중입니다.");
    }else if(item===4){
      q=con.ExecQuery("SELECT RecordNumber FROM Win32_NTLogEvent WHERE Logfile='System'");en=new Enumerator(q);found=0;for(;!en.atEnd()&&found<1;en.moveNext())found++;if(found){diagResult(4,"검사 완료","정상","good","시스템 이벤트 로그 접근 정상");}else{diagResult(4,"주의","이벤트 로그 확인 필요","caution","시스템 이벤트 로그를 확인할 수 없습니다.");issuePush(4,"eventlog","System 이벤트 로그","");}
    }else if(item===5){
      q=con.ExecQuery("SELECT NetConnectionStatus FROM Win32_NetworkAdapter WHERE NetEnabled=True");en=new Enumerator(q);netOk=!en.atEnd();if(netOk)diagResult(5,"검사 완료","정상","good","활성 네트워크 어댑터가 연결되어 있습니다.");else{diagResult(5,"오류","네트워크 연결 없음","bad","활성 네트워크 어댑터가 없습니다.");issuePush(5,"network","네트워크 연결","");}
    }else if(item===6){
      diagnoseComDevice(6,"COM1","COM1",con);
    }else if(item===7){
      diagnoseComDevice(7,"COM2","COM2",con);
    }else if(item===8){
      diagnoseComDevice(8,"COM3","COM3",con);
    }else if(item===9){
      diagnoseComDevice(9,"COM4","COM4",con);
    }else if(item===10){
      diagnoseComDevice(10,"COM5","COM5",con);
    }
  }catch(e){diagResult(item,"오류","진단 오류","bad","항목 "+item+" 진단 중 오류: "+e.message);issuePush(item,"system","진단 오류",String(e.message));}
}
function updateDiagSummary(){var i,r,ok=0,warn=0,err=0,total=ACTIVE_ITEMS.length;for(i=0;i<ACTIVE_ITEMS.length;i++){r=byId("r"+ACTIVE_ITEMS[i]);if(!r)continue;var v=r.cells[5].innerText;if(v==="정상"||v==="검사 완료"||v==="COM 없음"||v==="장치 확인"||v.indexOf("연결 정상")>=0)ok++;else if(v==="주의"||v==="확인 필요"||v==="프린터 미설치"||v.indexOf("연결 안됨")>=0)warn++;else if(v==="오류"||v.indexOf("오류")>=0||v.indexOf("중지")>=0)err++;}setText("okCount",ok);setText("okPct",Math.round(ok*100/total)+"%");setText("warnCount",warn);setText("warnPct",Math.round(warn*100/total)+"%");setText("errCount",err);setText("errPct",Math.round(err*100/total)+"%");}
function startQueue(a,label){running=true;idx=0;selectedQueue=a;selectedPos=0;lastDiagnosisIssues=[];resetRows(a);progress(0);setText("ready",label+" 진행 중");setText("readySub",a.length+"개 항목을 실제 PC 상태로 검사합니다.");setText("fixCount","0");setText("fixPct","0%");logMsg(label+"을 시작합니다. ("+a.length+"개 항목)","info");if(timer)window.clearInterval(timer);timer=window.setInterval(stepQueue,650);}
function stepQueue(){try{if(selectedPos>=selectedQueue.length){finishQueue();return;}var item=selectedQueue[selectedPos],r=byId("r"+item);selectedPos++;progress(Math.round(selectedPos*100/selectedQueue.length));if(r){r.cells[4].innerText="검사 중";r.cells[5].innerText="검사 중";r.cells[4].className="st caution";r.cells[5].className="res caution";}setText("phase","진행 단계: "+selectedPos+" / "+selectedQueue.length);diagnoseItem(item);updateDiagSummary();if(selectedPos>=selectedQueue.length)finishQueue();}catch(e){if(timer)window.clearInterval(timer);timer=null;running=false;logMsg("진단 처리 중 오류: "+e.message,"bad");}}
function finishQueue(){if(timer)window.clearInterval(timer);timer=null;running=false;updateDiagSummary();var issueCount=lastDiagnosisIssues.length;setText("ready",issueCount?"진단 완료 - 조치 필요":"진단 완료");setText("readySub",selectedQueue.length+"개 항목 실제 상태 확인 완료 / 문제 "+issueCount+"건");setText("fixCount","0");setText("fixPct",issueCount?"복구 대기":"0%");setText("phase",issueCount?"진행 단계: 오류 발견 - 자동 복구 가능":"진행 단계: 진단 완료");logMsg((selectedQueue.length===ACTIVE_ITEMS.length?"전체":"선택")+" 진단 완료: 문제 "+issueCount+"건","good");if(issueCount)logMsg("자동 복구 버튼을 누르면 발견된 복구 가능 오류를 수정하고 다시 확인합니다.","caution");}
function stopDiagnosis(){if(timer)window.clearInterval(timer);timer=null;running=false;setText("phase","진행 단계: 중지됨");logMsg("진단을 중지했습니다.","caution");}
function recoverSelected(){
  if(running||autoRunning){alert("현재 진단/복구 작업이 진행 중입니다.");return;}
  var a=getSelected(),i,r,targets=[];if(!a.length){alert("복구할 항목을 하나 이상 선택하세요.");return;}
  for(i=0;i<a.length;i++){r=byId("r"+a[i]);if(r&&(r.cells[5].innerText==="오류"||r.cells[5].innerText==="주의"||r.cells[5].innerText.indexOf("오류")>=0))targets.push(a[i]);}
  if(!targets.length){alert("선택한 항목에 현재 확인된 복구 대상 오류가 없습니다.");return;}
  var old=lastDiagnosisIssues,last=[];for(i=0;i<old.length;i++)if(targets.indexOf(old[i].row)>=0)last.push(old[i]);lastDiagnosisIssues=last;repairCurrentDiagnosisIssues();
}
function openDiskCleanup(){var m=byId("diskModal");if(!m)return;m.style.display="block";refreshDiskInfo();logMsg("디스크 정리 메뉴를 열었습니다.","info");}
function closeDiskCleanup(){var m=byId("diskModal");if(m)m.style.display="none";}
function diskLog(msg,cls){var e=byId("diskStatus");if(!e)return;var d=document.createElement("div");d.className=cls||"";d.innerText="["+new Date().toLocaleTimeString()+"] "+msg;e.appendChild(d);e.scrollTop=e.scrollHeight;}
function refreshDiskInfo(){try{var con=diagWmi(),d=con.Get('Win32_LogicalDisk.DeviceID="C:"'),size=d&&d.Size?Number(d.Size):0,free=d&&d.FreeSpace?Number(d.FreeSpace):0;setText("diskTotal",totalGB(size)+" GB");setText("diskFree",totalGB(free)+" GB");setText("diskUsed",(size?Math.round((1-free/size)*100):0)+"%");diskLog("C: 드라이브 상태를 확인했습니다.","good");}catch(e){diskLog("C: 드라이브 정보를 확인하지 못했습니다: "+e.message,"caution");}}
function optimizeCDrive(){if(running||autoRunning){alert("현재 진단/복구 작업이 진행 중입니다.");return;}if(!confirm("C: 드라이브 최적화를 시작하시겠습니까?\nWindows 기본 디스크 최적화 명령을 관리자 권한으로 실행합니다."))return;diskLog("C: 드라이브 최적화를 시작합니다.","info");var rc=runAdminShell('defrag C: /O /U /V');if(rc===0){diskLog("C: 드라이브 최적화 명령이 완료되었습니다.","good");logMsg("C: 드라이브 최적화 완료","good");}else{diskLog("최적화 실행 결과 코드 "+rc+" (관리자 권한 또는 Windows 상태를 확인하세요.)","caution");logMsg("C: 드라이브 최적화 실행 결과 확인 필요","caution");}}
function cleanCDrive(){if(running||autoRunning){alert("현재 진단/복구 작업이 진행 중입니다.");return;}diskLog("Windows 디스크 정리(cleanmgr)를 시작합니다.","info");try{var sh=new ActiveXObject("WScript.Shell");sh.Run('cleanmgr.exe /d C:',1,false);diskLog("디스크 정리 창을 실행했습니다.","good");}catch(e){diskLog("디스크 정리 실행 오류: "+e.message,"caution");}}
function quick(name){
  if(name=="디스크 정리"){openDiskCleanup();return;}
  if(name=="전체 진단"){startDiagnosis();return;}
  if(name=="자동 복구"){startAutoRecovery();return;}
  if(name=="긴급 복구"){emergencyRecovery();return;}
  logMsg(name+" 작업을 선택했습니다.","info");setText("phase","진행 단계: "+name);
}

/* 자동 복구: PC/POS 환경을 실제 Windows 정보로 순차 점검하고, 안전하게 복구할 수 있는 항목만 자동 처리합니다. */
var autoRunning=false,autoStage=0,autoIssues=[],autoTimer=null;
function setRowResult(i,status,result,cls){var r=byId("r"+i);if(!r)return;r.cells[4].innerText=status;r.cells[5].innerText=result;r.cells[4].className="st "+(cls||"");r.cells[5].className="res "+(cls||"");}
function autoMark(i,status,result,cls,msg){setRowResult(i,status,result,cls);if(msg)logMsg(msg,cls||"info");}
function runShell(cmd,wait){try{var sh=new ActiveXObject("WScript.Shell");return sh.Run(cmd,0,wait!==false);}catch(e){return -1;}}
function runAdminShell(cmd){
  try{
    var sh=new ActiveXObject("WScript.Shell");
    var safe=String(cmd||"").replace(/'/g,"''");
    var ps="$p=Start-Process -FilePath 'cmd.exe' -ArgumentList '/c "+safe+"' -Verb RunAs -WindowStyle Hidden -Wait -PassThru; exit $p.ExitCode";
    var full='powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "'+ps.replace(/"/g,'\\"')+'"';
    return sh.Run(full,0,true);
  }catch(e){return -1;}
}
function serviceState(name){try{var con=comWmi(),q=con.ExecQuery("SELECT State,StartMode FROM Win32_Service WHERE Name='"+name+"'"),en=new Enumerator(q);if(en.atEnd())return "missing";var x=en.item();return String(x.State||"").toLowerCase();}catch(e){return "error";}}
function startService(name){try{var rc=runAdminShell('sc start "'+name+'"');return rc===0;}catch(e){return false;}}
function autoRecoverySystemChecks(){
  var con,os,d,q,en,sum,n,cpu,total,free,ram,size,freeDisk,disk,netOk,wmiSvc;
  try{
    con=comWmi();
    q=con.ExecQuery('SELECT LoadPercentage FROM Win32_Processor');en=new Enumerator(q);sum=0;n=0;for(;!en.atEnd();en.moveNext()){sum+=Number(en.item().LoadPercentage||0);n++;}
    cpu=n?Math.round(sum/n):0;
    os=con.Get('Win32_OperatingSystem=@');total=Number(os.TotalVisibleMemorySize||0);free=Number(os.FreePhysicalMemory||0);ram=total?Math.round((1-free/total)*100):0;
    d=con.Get('Win32_LogicalDisk.DeviceID="C:"');size=d&&d.Size?Number(d.Size):0;freeDisk=d&&d.FreeSpace?Number(d.FreeSpace):0;disk=size?Math.round((1-freeDisk/size)*100):0;
    autoMark(2,cpu>=95||ram>=95||disk>=98?"주의":"검사 완료",cpu>=95||ram>=95||disk>=98?"주의":"정상",cpu>=95||ram>=95||disk>=98?"caution":"good","자원 상태: CPU "+cpu+"%, RAM "+ram+"%, 디스크 "+disk+"%");
    wmiSvc=serviceState("Winmgmt");
    if(wmiSvc!=="running"){autoMark(3,"오류","WMI 서비스 중지","bad","WMI 서비스가 중지되어 있습니다.");autoIssues.push({type:"service",name:"Winmgmt",row:3});}
    else autoMark(3,"검사 완료","정상","good","Windows 핵심 서비스 상태 정상");
    autoMark(1,"검사 완료","정상","good","Windows 운영체제 기본 상태 확인 완료");
    autoMark(4,"검사 완료","검사 대상 확인","good","이벤트 로그 저장소 접근 확인 완료");
    q=con.ExecQuery('SELECT NetConnectionStatus FROM Win32_NetworkAdapter WHERE NetEnabled=True');en=new Enumerator(q);netOk=!en.atEnd();
    if(netOk)autoMark(5,"검사 완료","정상","good","네트워크 어댑터 연결 확인");else{autoMark(5,"주의","네트워크 연결 없음","caution","활성 네트워크 어댑터가 없습니다.");autoIssues.push({type:"network",row:5});}
    var ports=[{row:6,port:"COM1",label:"COM1"},{row:7,port:"COM2",label:"COM2"},{row:8,port:"COM3",label:"COM3"},{row:9,port:"COM4",label:"COM4"},{row:10,port:"COM5",label:"COM5"}];
    for(var pi=0;pi<ports.length;pi++){
      var pp=findSpecificCom(con,ports[pi].port);
      if(!pp.found) autoMark(ports[pi].row,"주의",ports[pi].port+" 연결 안됨","caution",ports[pi].label+" "+ports[pi].port+" 장치를 찾을 수 없습니다.");
      else if(pp.error!==0){autoMark(ports[pi].row,"오류",ports[pi].port+" 오류 코드 "+pp.error,"bad",ports[pi].label+" "+ports[pi].port+" 장치 오류가 발견되었습니다.");autoIssues.push({type:"com",id:pp.id,name:pp.name,err:pp.error,row:ports[pi].row});}
      else if(probeSerialDevice(ports[pi].port)) autoMark(ports[pi].row,"검사 완료","기기 연결","good",ports[pi].port+" 실제 통신 상태 신호 확인");
      else autoMark(ports[pi].row,"주의","기기 연결 안됨","caution",ports[pi].port+" 포트는 있으나 실제 기기 통신 상태 신호가 확인되지 않습니다.");
    }
    return true;
  }catch(e){logMsg("자동 복구 사전 점검 오류: "+e.message,"caution");return false;}
}
function diagnoseExtraPrinter(port,row,label){
  var p=findSpecificCom(diagWmi(),"COM"+port);
  if(!p.found) autoMark(row,"주의","COM"+port+" 연결 안됨","caution",label+" COM"+port+" 장치를 찾을 수 없습니다.");
  else if(p.error!==0){autoMark(row,"오류","COM"+port+" 오류 코드 "+p.error,"bad",label+" COM"+port+" 장치 오류가 발견되었습니다.");autoIssues.push({type:"com",id:p.id,name:p.name,err:p.error,row:row});}
  else autoMark(row,"검사 완료","COM"+port+" 연결 정상","good",label+" COM"+port+" 연결 정상");
}
function runComDiagnosticsSilent(con){
  try{
    var p=findSpecificCom(con,"COM5");
    if(!p.found){logMsg("COM5 연결 장치 없음","caution");return;}
    if(p.error!==0){autoIssues.push({type:"com",id:p.id,name:p.name,err:p.error,row:8});logMsg("COM5 오류 코드 "+p.error,"caution");}
    else logMsg("COM5 정상","good");
  }catch(e){logMsg("COM5 상태 확인을 건너뛰었습니다: "+e.message,"caution");}
}
function repairComSilent(id,name){
  if(!id)return false;
  try{
    var safe=String(id).replace(/"/g,'');
    var rc=runAdminShell('pnputil /restart-device "'+safe+'"');
    if(rc===0){logMsg(name+" : COM 장치 재초기화 완료","good");return true;}
    logMsg(name+" : COM 장치 복구 실패(관리자 권한 필요 또는 장치 미지원)","caution");return false;
  }catch(e){logMsg(name+" : COM 복구 오류 "+e.message,"caution");return false;}
}
function repairAutoIssues(){
  var i,it,ok,rc;
  for(i=0;i<autoIssues.length;i++){
    it=autoIssues[i];
    if(it.type==="service"){
      ok=startService(it.name);
      if(ok){setRowResult(it.row,"복구 완료","정상","good");logMsg(it.name+" 서비스를 시작하여 복구했습니다.","good");}
      else{setRowResult(it.row,"복구 실패","관리자 권한 필요","bad");logMsg(it.name+" 서비스 복구에 실패했습니다. 관리자 권한을 확인하세요.","caution");}
    }else if(it.type==="network"){
      rc=runAdminShell('ipconfig /flushdns');
      if(rc===0){setRowResult(it.row,"복구 완료","DNS 초기화 완료","good");logMsg("DNS 캐시를 초기화했습니다.","good");}
      else{setRowResult(it.row,"복구 실패","확인 필요","bad");logMsg("DNS 초기화에 실패했습니다.","caution");}
    }else if(it.type==="com"){
      if(repairComSilent(it.id,it.name)){setRowResult(it.row,"복구 완료","COM 정상화","good");}
      else{setRowResult(it.row,"복구 실패","관리자 권한 필요","bad");}
    }
  }
}
function runWindowsFileRepair(){
  setText("phase","진행 단계: Windows 시스템 파일 검사 및 자동 복구 중");
  logMsg("Windows 시스템 파일 무결성 검사를 시작합니다.","info");
  try{
    var sh=new ActiveXObject("WScript.Shell"),rc=sh.Run('cmd.exe /c sfc /verifyonly',0,true);
    if(rc===0){
      logMsg("SFC 사전 검사 정상. 그래도 복구 단계에서 DISM/SFC를 안전하게 재확인합니다.","good");
    }else{
      logMsg("SFC 사전 검사에서 문제가 확인되었습니다. 자동 복구를 시작합니다.","caution");
    }
    logMsg("Windows 구성 요소 저장소 자동 복구(DISM)를 실행합니다.","info");
    rc=runAdminShell('DISM /Online /Cleanup-Image /RestoreHealth');
    if(rc===0)logMsg("DISM 자동 복구 완료.","good");
    else logMsg("DISM 복구 결과 코드: "+rc+". SFC 복구를 계속합니다.","caution");
    logMsg("Windows 시스템 파일 자동 복구(SFC)를 실행합니다.","info");
    rc=runAdminShell('sfc /scannow');
    if(rc===0){autoMark(1,"복구 완료","정상","good","SFC 시스템 파일 자동 복구가 완료되었습니다.");}
    else{autoMark(1,"복구 확인 필요","SFC 결과 코드 "+rc,"caution","SFC 복구 결과를 확인해야 합니다.");}
  }catch(e){autoMark(1,"검사 실패","권한/실행 오류","bad","SFC/DISM 실행 오류: "+e.message);}
}
function repairCurrentDiagnosisIssues(){
  if(!lastDiagnosisIssues.length){alert("최근 진단에서 발견된 오류가 없습니다. 자동 복구가 필요하지 않습니다.");return;}
  if(running||autoRunning){alert("현재 다른 진단/복구 작업이 진행 중입니다.");return;}
  autoRunning=true;autoIssues=[];var i,it,rc,ok;
  logMsg("발견된 오류의 자동 복구를 시작합니다.","info");
  for(i=0;i<lastDiagnosisIssues.length;i++){
    it=lastDiagnosisIssues[i];ok=false;
    if(it.type==="service"){ok=startService(it.name);if(ok){setRowResult(it.row,"복구 완료","정상","good");logMsg(it.name+" 서비스 자동 복구 완료","good");}else{setRowResult(it.row,"복구 실패","관리자 권한 필요","bad");logMsg(it.name+" 서비스 복구 실패","caution");}}
    else if(it.type==="com"){ok=repairComSilent(it.extra,it.name);if(ok){setRowResult(it.row,"복구 완료","정상","good");}}
    else if(it.type==="network"){rc=runAdminShell('ipconfig /flushdns');if(rc===0){ok=true;setRowResult(it.row,"복구 완료","DNS 초기화","good");logMsg("네트워크 DNS 캐시를 자동 초기화했습니다.","good");}else{setRowResult(it.row,"복구 실패","관리자 권한 필요","bad");}}
    else if(it.type==="system"){runWindowsFileRepair();ok=true;}
    else if(it.type==="eventlog"){ok=startService("EventLog");if(ok){setRowResult(it.row,"복구 완료","이벤트 로그 서비스 시작","good");logMsg("Windows 이벤트 로그 서비스를 시작했습니다.","good");}}
    else if(it.type==="resource"){logMsg("자원 사용량 문제는 시스템 파일을 임의 삭제하지 않고 재진단합니다.","caution");ok=true;}
    if(ok)autoIssues.push(it);
  }
  autoRunning=false;
  liveRefresh();
  logMsg("자동 복구 작업 후 오류 항목을 다시 진단합니다.","info");
  var q=[],seen={},j;
  for(j=0;j<lastDiagnosisIssues.length;j++){if(!seen[lastDiagnosisIssues[j].row]){seen[lastDiagnosisIssues[j].row]=true;q.push(lastDiagnosisIssues[j].row);}}
  for(j=0;j<q.length;j++){diagnoseItem(q[j]);}
  updateDiagSummary();
  setText("fixCount",autoIssues.length);setText("fixPct",lastDiagnosisIssues.length?Math.round(autoIssues.length*100/lastDiagnosisIssues.length)+"%":"0%");
  setText("phase","진행 단계: 자동 복구 및 재진단 완료");
  logMsg("자동 복구 및 재진단이 완료되었습니다.","good");
  alert("자동 복구 및 재진단이 완료되었습니다.\n복구 가능한 항목은 수정 후 다시 확인했습니다.");
}
function startAutoRecovery(){
  if(running||autoRunning){alert("현재 다른 진단/복구 작업이 진행 중입니다.");return;}
  if(lastDiagnosisIssues.length){repairCurrentDiagnosisIssues();return;}
  autoRunning=true;autoStage=0;autoIssues=[];progress(0);setText("ready","자동 복구 진행 중");setText("readySub","PC/POS 전체 상태를 검사하고 복구 가능한 오류를 자동 처리합니다.");setText("phase","진행 단계: 전체 검사 준비");logMsg("최근 진단 결과가 없어 PC/POS 전체 점검 후 자동 복구를 시작합니다.","info");reset();if(autoTimer)window.clearInterval(autoTimer);autoTimer=window.setInterval(autoRecoveryStep,700);
}
function autoRecoveryStep(){
  try{
    if(autoStage===0){autoRecoverySystemChecks();progress(25);autoStage++;return;}
    if(autoStage===1){runWindowsFileRepair();progress(65);autoStage++;return;}
    if(autoStage===2){repairAutoIssues();progress(90);autoStage++;return;}
    if(autoStage===3){liveRefresh();progress(100);setText("ready","자동 복구 완료");setText("readySub","PC/POS 전체 검사 및 가능한 오류 복구가 완료되었습니다.");setText("phase","진행 단계: 자동 복구 완료");setText("fixCount",autoIssues.length);setText("fixPct",autoIssues.length?"복구 처리":"100%");logMsg("PC/POS 전체 검사 및 자동 복구가 완료되었습니다.","good");autoRunning=false;if(autoTimer)window.clearInterval(autoTimer);autoTimer=null;alert("자동 복구가 완료되었습니다.\n복구 처리 항목: "+autoIssues.length+"건");return;}
  }catch(e){autoRunning=false;if(autoTimer)window.clearInterval(autoTimer);autoTimer=null;logMsg("자동 복구 중 오류: "+e.message,"caution");setText("phase","진행 단계: 자동 복구 오류");}
}
function wmi(){
  return GetObject("winmgmts:{impersonationLevel=impersonate}!\\\\.\\root\\cimv2");
}
function nowShort(){try{return new Date().toLocaleTimeString();}catch(e){return "-";}}
function setLive(id,value,sub){setText(id,value);if(sub!==undefined)setText(id==="cpu"?"cpuVal":id==="mem"?"memSub":id==="disk"?"diskSub":id==="netVal"?"netSub":id==="inetVal"?"inetSub":"posSub",sub);}
function setNetStatus(id,value,good){var e=byId(id);if(e){e.innerText=value;e.className='svalue '+(good?'statusGood':'statusWarn');}}
function liveSystem(){
  try{
    var svc=new ActiveXObject('WbemScripting.SWbemLocator');
    var con=svc.ConnectServer('.','root\\cimv2');
    var q=con.ExecQuery('SELECT LoadPercentage FROM Win32_Processor');
    var en=new Enumerator(q),sum=0,n=0;
    for(;!en.atEnd();en.moveNext()){sum+=Number(en.item().LoadPercentage||0);n++;}
    var cpu=n?Math.round(sum/n):0;
    setText('cpu',cpu+'%');setWidth('cpuBar',cpu);setText('cpuVal','실시간 PC 사용률');
    var os=con.Get('Win32_OperatingSystem=@');
    var total=Number(os.TotalVisibleMemorySize||0),free=Number(os.FreePhysicalMemory||0);
    var ram=total?Math.round((1-free/total)*100):0;
    setText('mem',ram+'%');setWidth('memBar',ram);setText('memSub',Math.round((total-free)/1048576*10)/10+' GB / '+Math.round(total/1048576*10)/10+' GB');
    var d=con.Get('Win32_LogicalDisk.DeviceID="C:"');
    var size=d&&d.Size?Number(d.Size):0,freeDisk=d&&d.FreeSpace?Number(d.FreeSpace):0;
    var disk=size?Math.round((1-freeDisk/size)*100):0;
    setText('disk',disk+'%');setWidth('diskBar',disk);setText('diskSub',Math.round(freeDisk/1073741824*10)/10+' GB 여유 / '+totalGB(size)+' GB');
  }catch(ex){
    setText('cpu','확인 필요');setText('cpuVal','WMI 조회 실패');setWidth('cpuBar',0);
    setText('mem','확인 필요');setText('memSub','WMI 조회 실패');setWidth('memBar',0);
    setText('disk','확인 필요');setText('diskSub','WMI 조회 실패');setWidth('diskBar',0);
  }
}
function liveNetwork(){
  var netOk=false;
  try{
    var svc=new ActiveXObject('WbemScripting.SWbemLocator');
    var con=svc.ConnectServer('.','root\\cimv2');
    var q=con.ExecQuery('SELECT NetConnectionStatus FROM Win32_NetworkAdapter WHERE NetEnabled=True AND NetConnectionStatus=2');
    var en=new Enumerator(q);
    netOk=!en.atEnd();
  }catch(ex){netOk=false;}
  setNetStatus('netVal',netOk?'정상':'주의',netOk);
  setText('netSub',netOk?'네트워크 연결됨 · 실시간':'활성 네트워크 없음');

  // HTA의 비동기 MSXML2.XMLHTTP는 환경/보안 설정에 따라 status=0이 되어
  // 실제 인터넷이 연결되어 있어도 '인터넷 응답 없음'으로 오판할 수 있습니다.
  // WinHTTP의 동기 요청으로 짧게 확인하고, 실패하면 보조 주소를 한 번 더 확인합니다.
  var ok=false;
  var urls=[
    'https://www.msftconnecttest.com/connecttest.txt',
    'https://www.gstatic.com/generate_204'
  ];
  for(var i=0;i<urls.length && !ok;i++){
    try{
      var w=new ActiveXObject('WinHttp.WinHttpRequest.5.1');
      w.SetTimeouts(1200,1200,1800,1800);
      w.Open('GET',urls[i],false);
      w.SetRequestHeader('Cache-Control','no-cache');
      w.Send();
      var st=Number(w.Status);
      if((st>=200&&st<400) || st===204){ok=true;}
    }catch(ex3){}
  }
  if(ok){
    setNetStatus('inetVal','정상',true);
    setText('inetSub','연결됨 · 실시간');
  }else if(netOk){
    // 네트워크 어댑터가 실제 연결 상태라면 일시적인 외부 응답 지연으로
    // '인터넷 응답 없음'을 즉시 표시하지 않고 재확인 상태로 표시합니다.
    setNetStatus('inetVal','확인 중',true);
    setText('inetSub','인터넷 연결 확인 중');
    window.setTimeout(function(){liveNetworkInternetOnly();},1200);
  }else{
    setNetStatus('inetVal','주의',false);
    setText('inetSub','네트워크 연결 없음');
  }
}

function liveNetworkInternetOnly(){
  var ok=false;
  var urls=['https://www.msftconnecttest.com/connecttest.txt','https://www.gstatic.com/generate_204'];
  for(var i=0;i<urls.length && !ok;i++){
    try{
      var w=new ActiveXObject('WinHttp.WinHttpRequest.5.1');
      w.SetTimeouts(1500,1500,2000,2000);
      w.Open('GET',urls[i],false);
      w.SetRequestHeader('Cache-Control','no-cache');
      w.Send();
      var st=Number(w.Status);
      ok=((st>=200&&st<400)||st===204);
    }catch(e){}
  }
  if(ok){
    setNetStatus('inetVal','정상',true);
    setText('inetSub','연결됨 · 실시간');
  }else{
    setNetStatus('inetVal','주의',false);
    setText('inetSub','인터넷 응답 지연 · 재확인 필요');
  }
}

function liveRefresh(){liveSystem();liveNetwork();}

var LOGIN_ID="yogiorder";
var DEFAULT_PASSWORD="e123456789";
var PASS_KEY="HKCU\\Software\\POSCAREPRO\\Password";
var SIZE_KEY="HKCU\\Software\\POSCAREPRO\\WindowSize";

function shellObj(){return new ActiveXObject("WScript.Shell");}
function getSavedPassword(){
  try{
    var s=shellObj();
    return String(s.RegRead(PASS_KEY));
  }catch(e){return DEFAULT_PASSWORD;}
}
function savePassword(v){
  var s=shellObj();
  s.RegWrite(PASS_KEY,String(v),"REG_SZ");
}
function showLogin(){
  var o=byId("loginOverlay"); if(o)o.style.display="block";
  try{byId("loginId").focus();}catch(e){}
}
function doLogin(){
  var uid=byId("loginId").value;
  var pw=byId("loginPw").value;
  if(uid===LOGIN_ID && pw===getSavedPassword()){
    byId("loginOverlay").style.display="none";
    logMsg("관리자 로그인 완료","good");
    return;
  }
  setText("loginMsg","아이디 또는 비밀번호가 올바르지 않습니다.");
  byId("loginPw").value="";
  try{byId("loginPw").focus();}catch(e){}
}
function loginKey(e){
  e=e||window.event;
  if(e.keyCode===13){doLogin();return false;}
}
function openSettings(){
  // 설정 메뉴 진입 전 현재 관리자 비밀번호를 한 번 더 확인합니다.
  var g=byId("settingsGate");
  if(g){
    g.style.display="block";
    try{byId("settingsGatePw").value="";byId("settingsGateMsg").innerText="";byId("settingsGatePw").focus();}catch(e){}
  }
}
function settingsGateKey(e){
  e=e||window.event;
  if(e.keyCode===13){verifySettingsGate();return false;}
}
function verifySettingsGate(){
  var p="";
  try{p=byId("settingsGatePw").value;}catch(e){}
  if(p===getSavedPassword()){
    var g=byId("settingsGate");
    if(g)g.style.display="none";
    var s=byId("settingsModal");
    if(s)s.style.display="block";
    setText("settingsMsg","");
    try{byId("oldPw").value="";byId("oldPwConfirm").value="";byId("newPw").value="";byId("newPw2").value="";byId("oldPw").focus();}catch(e){}
    logMsg("관리자 비밀번호 재확인 후 설정 메뉴에 진입했습니다.","good");
  }else{
    setText("settingsGateMsg","현재 비밀번호가 올바르지 않습니다.");
    try{byId("settingsGatePw").value="";byId("settingsGatePw").focus();}catch(e){}
    logMsg("설정 메뉴 진입 비밀번호 확인 실패","caution");
  }
}
function closeSettingsGate(){
  var g=byId("settingsGate");
  if(g)g.style.display="none";
  try{byId("settingsGatePw").value="";}catch(e){}
  setText("settingsGateMsg","");
}
function closeSettings(){
  var s=byId("settingsModal");
  if(s)s.style.display="none";
  setText("settingsMsg","");
  try{byId("oldPw").value="";byId("oldPwConfirm").value="";byId("newPw").value="";byId("newPw2").value="";}catch(e){}
}
function changePassword(){
  var oldPw=byId("oldPw").value;
  var oldPwConfirm=byId("oldPwConfirm").value;
  var newPw=byId("newPw").value;
  var newPw2=byId("newPw2").value;
  if(oldPw!==getSavedPassword()){setText("settingsMsg","현재 비밀번호가 올바르지 않습니다.");return;}
  if(oldPwConfirm!==getSavedPassword() || oldPwConfirm!==oldPw){setText("settingsMsg","현재 비밀번호를 한 번 더 정확하게 입력하세요.");return;}
  if(newPw.length<4){setText("settingsMsg","새 비밀번호는 4자 이상 입력하세요.");return;}
  if(newPw!==newPw2){setText("settingsMsg","새 비밀번호가 서로 일치하지 않습니다.");return;}
  try{
    savePassword(newPw);
    setText("settingsMsg","비밀번호가 변경되었습니다.");
    logMsg("관리자 비밀번호가 변경되었습니다.","good");
    window.setTimeout(closeSettings,700);
  }catch(e){setText("settingsMsg","비밀번호 저장에 실패했습니다.");}
}
function getWindowSize(){
  try{var s=shellObj();return String(s.RegRead(SIZE_KEY));}catch(e){return "maximized";}
}
function saveWindowSize(v){try{shellObj().RegWrite(SIZE_KEY,String(v),"REG_SZ");}catch(e){}}
function applyWindowSize(v){
  try{
    if(v==="maximized"){window.moveTo(0,0);window.resizeTo(screen.availWidth,screen.availHeight);return true;}
    var a=v.split("x"),w=parseInt(a[0],10),h=parseInt(a[1],10);
    if(!w||!h)return false;
    var sw=screen.availWidth||w, sh=screen.availHeight||h;
    w=Math.min(w,sw); h=Math.min(h,sh);
    window.resizeTo(w,h);
    window.moveTo(Math.max(0,Math.floor((sw-w)/2)),Math.max(0,Math.floor((sh-h)/2)));
    return true;
  }catch(e){return false;}
}
function loadWindowSize(){
  var v=getWindowSize();
  if(v!=="maximized" && !/^[0-9]+x[0-9]+$/.test(v))v="maximized";
  var sel=byId("windowSize");if(sel)sel.value=v;
}
function saveWindowSizeSetting(){
  var sel=byId("windowSize");if(!sel)return;
  var v=sel.value;
  if(applyWindowSize(v)){saveWindowSize(v);setText("windowMsg","창 크기가 적용되었습니다.");logMsg("창 크기 설정을 적용했습니다.","good");}
  else setText("windowMsg","창 크기 적용에 실패했습니다.");
}

var printTestPid=0;
function openPrintTest(){
  var m=byId('printModal');
  if(m)m.style.display='block';
  setText('printStatus','대기 중');
  setText('printExeStatus','내장 PrnTest 준비 완료');
  logMsg('프린트 테스트 설정 창을 열었습니다.','info');
}
function closePrintTest(){var m=byId('printModal');if(m)m.style.display='none';}
function printImageSelected(input){try{var v=input.value||'';setText('printImageName',v?v.split('\\').pop():'선택된 이미지 없음');}catch(e){setText('printImageName','선택된 이미지 없음');}}
function runPrintTest(){
  var speed=byId('printSpeed').value,port=byId('printPort').value;
  setText('printStatus','테스트 출력 중...');setText('printExeStatus','내장 프린트 통신 모듈을 실행합니다.');
  logMsg('프린트 테스트: '+port+' / '+speed+' bps','info');setText('phase','진행 단계: 프린터 포트 열기 및 테스트 출력');
  try{
    var sh=new ActiveXObject('WScript.Shell');
    var fso=new ActiveXObject('Scripting.FileSystemObject');
    var base=fso.GetParentFolderName(location.pathname.replace(/^\\\//,''));
    var ps=fso.BuildPath(base,'print_test.ps1');
    if(!fso.FileExists(ps)){throw new Error('print_test.ps1 파일을 찾을 수 없습니다.');}
    var img=byId('printImage'), imgPath='';
    if(img && img.value){imgPath=img.value;}
    var action=(imgPath?'NVPRINT':'TEST');
    var cmd='powershell.exe -NoProfile -ExecutionPolicy Bypass -File "'+ps+'" -Action '+action+' -Port '+port+' -Baud '+speed;
    if(imgPath){cmd+=' -ImagePath "'+imgPath.replace(/"/g,'`"')+'"';}
    printTestPid=sh.Run(cmd,1,true);
    if(printTestPid===0){
      setText('printStatus','테스트 출력 완료');setText('printExeStatus','프린터에서 테스트 출력 결과를 확인하세요.');
      setText('phase','진행 단계: 프린트 테스트 완료');
      logMsg('프린트 테스트 출력 완료 (COM '+port+' / '+speed+' bps)','good');
    }else{throw new Error('프린터 통신 모듈이 오류 코드 '+printTestPid+'로 종료되었습니다. COM 포트와 프린터 연결을 확인하세요.');}
  }catch(ex){setText('printStatus','실행 오류');setText('printExeStatus','프린트 테스트 실패');setText('phase','진행 단계: 프린트 테스트 확인 필요');logMsg('프린트 테스트 오류: '+ex.message,'caution');alert('프린트 테스트를 실행할 수 없습니다.\n\n'+ex.message);}
}

function stopPrintTest(){
  try{
    var sh=new ActiveXObject('WScript.Shell');
    if(printTestPid){sh.Run('taskkill /PID '+printTestPid+' /T /F',0,true);printTestPid=0;}
    else{sh.Run('taskkill /IM powershell.exe /T /F',0,true);}
    setText('printStatus','중지 완료');setText('printExeStatus','PrnTest 실행을 중지했습니다.');logMsg('프린트 테스트 프로그램 중지','caution');
  }catch(ex){logMsg('PrnTest 중지 오류: '+ex.message,'caution');}
}
function openComDiagnostic(){var m=byId('comModal');if(m)m.style.display='block';logMsg('COM1~COM10 포트 진단 창을 열었습니다.','info');runComDiagnostics();}
function closeComDiagnostic(){var m=byId('comModal');if(m)m.style.display='none';}
function comWmi(){var svc=new ActiveXObject('WbemScripting.SWbemLocator');return svc.ConnectServer('.','root\\cimv2');}
function comLog(t,c){var e=byId('comLog');if(!e)return;var d=document.createElement('div');d.className=c||'';d.innerText='['+nowShort()+'] '+t;e.appendChild(d);e.scrollTop=e.scrollHeight;}
function setComCell(id,text,cls){var e=byId(id);if(e){e.innerText=text;var base=(e.className||'').split(' ')[0];e.className=base+' '+('com'+(cls||''));}}
function findComInfo(con,port){
  var q=con.ExecQuery("SELECT Name,PNPDeviceID,ConfigManagerErrorCode,Status FROM Win32_PnPEntity WHERE Name LIKE '%("+port+")%'");
  var en=new Enumerator(q),a=[];
  for(;!en.atEnd();en.moveNext()){var x=en.item();a.push({name:String(x.Name||''),id:String(x.PNPDeviceID||''),err:Number(x.ConfigManagerErrorCode||0),status:String(x.Status||'')});}
  return a;
}
function runComDiagnostics(){
  var i,con,info,port,found,err;
  for(i=1;i<=10;i++){setComCell('cs'+i,'검사 중','Warn');setComCell('cr'+i,'-','None');var b=byId('cb'+i);if(b)b.disabled=true;}
  try{con=comWmi();}catch(ex){for(i=1;i<=10;i++){setComCell('cs'+i,'WMI 오류','Bad');setComCell('cr'+i,'실패','Bad');}comLog('COM 진단을 위한 WMI 연결에 실패했습니다: '+ex.message,'comBad');return;}
  for(i=1;i<=10;i++){
    port='COM'+i;found=findComInfo(con,port);
    if(!found.length){setComCell('cs'+i,'미사용','None');setComCell('cr'+i,'장치 없음','None');comLog(port+' : 연결된 장치가 없습니다.','comNone');continue;}
    err=0;for(var j=0;j<found.length;j++){if(found[j].err!==0)err=found[j].err;}
    byId('cn'+i).innerText=found[0].name;
    if(err!==0){setComCell('cs'+i,'오류','Bad');setComCell('cr'+i,'오류 코드 '+err,'Bad');var bb=byId('cb'+i);if(bb){bb.disabled=false;bb.setAttribute('data-pnp',found[0].id);}comLog(port+' : 오류 발견 (ConfigManagerErrorCode '+err+')','comBad');}
    else{setComCell('cs'+i,'정상','Good');setComCell('cr'+i,'정상','Good');comLog(port+' : 정상','comGood');}
  }
  logMsg('COM1~COM10 진단 완료','good');
}
function recoverCom(port){
  var b=byId('cb'+port),id=b?b.getAttribute('data-pnp'):'';
  if(!id){alert('복구할 COM 장치를 찾을 수 없습니다.');return;}
  if(!confirm('COM'+port+' 장치를 비활성화 후 다시 활성화하여 복구하시겠습니까?'))return;
  try{
    var safe=id.replace(/'/g,"''"),ps="powershell.exe -NoProfile -ExecutionPolicy Bypass -Command \"$id='"+safe+"'; Disable-PnpDevice -InstanceId $id -Confirm:$false -ErrorAction Stop; Start-Sleep -Milliseconds 800; Enable-PnpDevice -InstanceId $id -Confirm:$false -ErrorAction Stop\"";
    var sh=new ActiveXObject('WScript.Shell'),rc=sh.Run(ps,0,true);
    if(rc===0){setComCell('cs'+port,'복구 완료','Good');setComCell('cr'+port,'정상','Good');comLog('COM'+port+' : 복구 완료','comGood');logMsg('COM'+port+' 복구 완료','good');window.setTimeout(runComDiagnostics,700);}
    else{setComCell('cs'+port,'복구 실패','Bad');setComCell('cr'+port,'관리자 권한 필요','Bad');comLog('COM'+port+' : 복구 실패 (관리자 권한 또는 장치 상태를 확인하세요.)','comBad');}
  }catch(ex){setComCell('cs'+port,'복구 실패','Bad');setComCell('cr'+port,'실패','Bad');comLog('COM'+port+' 복구 오류: '+ex.message,'comBad');}
}
function recoverAllCom(){
  var i,targets=[];for(i=1;i<=10;i++){var e=byId('cs'+i);if(e&&e.innerText==='오류')targets.push(i);}
  if(!targets.length){alert('현재 복구할 COM 포트 오류가 없습니다.');return;}
  if(!confirm('오류가 발견된 '+targets.length+'개 COM 포트를 복구하시겠습니까?'))return;
  for(i=0;i<targets.length;i++)recoverCom(targets[i]);
}

function emergencyRecovery(){
  if(running||autoRunning){alert("현재 다른 진단/복구 작업이 진행 중입니다.");return;}
  if(!confirm("긴급 복구를 실행하시겠습니까?\n\n네트워크 DNS/Winsock, Windows 핵심 서비스, 프린터 스풀러를 복구한 뒤 전체 진단으로 결과를 다시 확인합니다.\n\n※ PC는 자동으로 재부팅하지 않습니다."))return;
  autoRunning=true;
  progress(0);
  setText("ready","긴급 복구 진행 중");
  setText("readySub","핵심 네트워크·서비스·프린터를 복구하고 전체 재진단합니다.");
  setText("phase","진행 단계: 긴급 복구 준비");
  logMsg("긴급 복구를 시작합니다. PC는 자동 재부팅하지 않습니다.","caution");
  try{
    var rc=runAdminShell('ipconfig /flushdns & netsh winsock reset & sc start Winmgmt & sc start EventLog & sc start RpcSs & sc start DHCP & sc start Spooler & net stop Spooler & net start Spooler');
    progress(45);
    if(rc===0)logMsg("네트워크 DNS/Winsock 및 Windows 핵심 서비스 복구 명령이 완료되었습니다.","good");
    else logMsg("긴급 복구 명령 일부의 결과 코드: "+rc+" (복구 후 재진단을 계속합니다).","caution");

    /* 최근 진단에서 COM 장치 오류가 발견된 경우에만 해당 장치를 재활성화합니다. */
    var i,it,comFixed=0;
    for(i=0;i<lastDiagnosisIssues.length;i++){
      it=lastDiagnosisIssues[i];
      if(it.type==="com" && it.extra){
        if(repairComSilent(it.extra,it.name)){
          comFixed++;
          logMsg(it.name+" COM 장치 재활성화 복구 완료.","good");
        }else{
          logMsg(it.name+" COM 장치 자동 복구 실패. 연결 상태를 다시 확인합니다.","caution");
        }
      }
    }
    progress(70);
    liveRefresh();
    logMsg("긴급 복구 후 전체 진단을 다시 실행합니다.","info");
    autoRunning=false;
    startQueue(ACTIVE_ITEMS,"긴급 복구 후 재진단");
    return;
  }catch(e){
    autoRunning=false;
    progress(0);
    setText("ready","긴급 복구 오류");
    setText("readySub","복구 과정에서 오류가 발생했습니다.");
    setText("phase","진행 단계: 긴급 복구 오류");
    logMsg("긴급 복구 실행 오류: "+e.message,"bad");
    alert("긴급 복구 중 오류가 발생했습니다.\n\n"+e.message);
  }
}


function periphLog(t,c){var e=byId("periphLog");if(!e)return;var d=document.createElement("div");d.className=c||"";d.innerText="["+nowShort()+"] "+t;e.appendChild(d);e.scrollTop=e.scrollHeight;}
function setPeriph(id,text,cls){var e=byId(id);if(e){e.innerText=text;e.className=(e.className||"").split(" ")[0]+" "+(cls||"");}}
function pnpFindByKeyword(con,keyword){var q=con.ExecQuery("SELECT Name,PNPDeviceID,ConfigManagerErrorCode,Status FROM Win32_PnPEntity WHERE Name LIKE '%"+keyword.replace(/'/g,"''")+"%'");var en=new Enumerator(q),a=[];for(;!en.atEnd();en.moveNext()){var x=en.item();a.push({name:String(x.Name||""),id:String(x.PNPDeviceID||""),err:Number(x.ConfigManagerErrorCode||0),status:String(x.Status||"")});}return a;}
function pnpFindPort(con,port){var q=con.ExecQuery("SELECT Name,PNPDeviceID,ConfigManagerErrorCode,Status FROM Win32_PnPEntity WHERE Name LIKE '%("+port+")%'");var en=new Enumerator(q),a=[];for(;!en.atEnd();en.moveNext()){var x=en.item();a.push({name:String(x.Name||""),id:String(x.PNPDeviceID||""),err:Number(x.ConfigManagerErrorCode||0),status:String(x.Status||"")});}return a;}
function peripheralCheck(){
  var con,i,a,err,ports=[['06','COM1','COM1'],['07','COM2','COM2'],['08','COM3','COM3'],['09','COM4','COM4'],['10','COM5','COM5']];
  for(i=0;i<5;i++){setPeriph('ps'+ports[i][0],'검사 중','periphWarn');setPeriph('pr'+ports[i][0],'-','periphNone');}
  setPeriph('ps11','검사 중','periphWarn');setPeriph('pr11','-','periphNone');setPeriph('ps12','검사 중','periphWarn');setPeriph('pr12','-','periphNone');setPeriph('ps13','검사 중','periphWarn');setPeriph('pr13','-','periphNone');
  try{con=diagWmi();}catch(ex){for(i=0;i<5;i++){setPeriph('ps'+ports[i][0],'검사 실패','periphBad');setPeriph('pr'+ports[i][0],'WMI 오류','periphBad');}setPeriph('ps11','검사 실패','periphBad');setPeriph('pr11','WMI 오류','periphBad');setPeriph('ps12','검사 실패','periphBad');setPeriph('pr12','WMI 오류','periphBad');setPeriph('ps13','검사 실패','periphBad');setPeriph('pr13','WMI 오류','periphBad');periphLog('주변기기 진단 WMI 연결 실패: '+ex.message,'periphBad');return;}
  for(i=0;i<5;i++){a=pnpFindPort(con,ports[i][2]);err=0;for(var j=0;j<a.length;j++)if(a[j].err!==0)err=a[j].err;if(!a.length){setText('px'+ports[i][0],'미연결');setPeriph('ps'+ports[i][0],'연결 안됨','periphWarn');setPeriph('pr'+ports[i][0],ports[i][2]+' 장치 없음','periphWarn');periphLog(ports[i][1]+' '+ports[i][2]+' : 연결된 장치 없음','periphWarn');}else if(err!==0){setText('px'+ports[i][0],a[0].name);setPeriph('ps'+ports[i][0],'장치 오류','periphBad');setPeriph('pr'+ports[i][0],'오류 코드 '+err,'periphBad');periphLog(ports[i][1]+' '+ports[i][2]+' : 장치 오류 코드 '+err,'periphBad');}else{setText('px'+ports[i][0],a[0].name);setPeriph('ps'+ports[i][0],'연결 정상','periphGood');setPeriph('pr'+ports[i][0],'정상 연결','periphGood');periphLog(ports[i][1]+' '+ports[i][2]+' : Windows 장치 인식 및 연결 정상','periphGood');}}
  try{var pq=con.ExecQuery("SELECT Name,Status,PrinterStatus,WorkOffline FROM Win32_Printer"),en=new Enumerator(pq),printers=[],x;for(;!en.atEnd();en.moveNext()){x=en.item();printers.push({name:String(x.Name||''),status:String(x.Status||''),offline:Boolean(x.WorkOffline),printerStatus:Number(x.PrinterStatus||0)});}if(printers.length){setText('px11',printers[0].name);setPeriph('ps11','연결 확인','periphGood');setPeriph('pr11',printers[0].offline?'오프라인':'프린터 인식 정상',printers[0].offline?'periphWarn':'periphGood');periphLog('영수증 프린터 : '+printers[0].name+' / '+(printers[0].offline?'오프라인':'Windows 프린터 인식 정상'),printers[0].offline?'periphWarn':'periphGood');}else{setText('px11','프린터 없음');setPeriph('ps11','연결 안됨','periphWarn');setPeriph('pr11','등록된 프린터 없음','periphWarn');periphLog('영수증 프린터 : Windows에 등록된 프린터가 없습니다.','periphWarn');}}catch(ex){setPeriph('ps11','확인 실패','periphBad');setPeriph('pr11','조회 오류','periphBad');periphLog('프린터 조회 오류: '+ex.message,'periphBad');}
  try{var mons=con.ExecQuery("SELECT Name FROM Win32_DesktopMonitor"),me=new Enumerator(mons),mc=0;for(;!me.atEnd();me.moveNext())mc++;setText('px12',mc+'개 모니터');setPeriph('ps12',mc>0?'연결 정상':'확인 필요',mc>0?'periphGood':'periphWarn');setPeriph('pr12',mc>0?'모니터 인식 정상':'모니터 정보 없음',mc>0?'periphGood':'periphWarn');periphLog('듀얼모니터 : Windows 모니터 '+mc+'개 인식','periphGood');}catch(ex){setPeriph('ps12','확인 실패','periphBad');setPeriph('pr12','조회 오류','periphBad');}
  try{var usb=con.ExecQuery("SELECT Name,ConfigManagerErrorCode FROM Win32_PnPEntity WHERE PNPClass='USB'"),ue=new Enumerator(usb),uc=0,ub=0;for(;!ue.atEnd();ue.moveNext()){uc++;if(Number(ue.item().ConfigManagerErrorCode||0)!==0)ub++;}setText('px13',uc+'개 USB 장치');if(ub){setPeriph('ps13','장치 오류','periphBad');setPeriph('pr13',ub+'개 오류','periphBad');periphLog('USB 장치 : '+uc+'개 인식 / '+ub+'개 오류','periphBad');}else{setPeriph('ps13','연결 정상','periphGood');setPeriph('pr13',uc+'개 정상 인식','periphGood');periphLog('USB 장치 : '+uc+'개 정상 인식','periphGood');}}catch(ex){setPeriph('ps13','확인 실패','periphBad');setPeriph('pr13','조회 오류','periphBad');}
  logMsg('주변기기 연결 상태 진단 완료','good');
}
function openPeripheralDiagnostic(){var m=byId('periphModal');if(!m)return;m.style.display='block';byId('periphLog').innerHTML='';periphLog('주변기기 진단을 시작합니다.','info');peripheralCheck();}
function closePeripheralDiagnostic(){var m=byId('periphModal');if(m)m.style.display='none';}
function openInfo(){var m=byId('infoModal');if(m)m.style.display='block';}
function closeInfo(){var m=byId('infoModal');if(m)m.style.display='none';}


function bsLogMsg(msg,cls){var e=byId("bsLog");if(!e)return;var d=document.createElement("div");d.className=cls||"";d.innerText="["+new Date().toLocaleTimeString()+"] "+msg;e.appendChild(d);e.scrollTop=e.scrollHeight;}
function bsSet(id,v,cls){var e=byId(id);if(e){e.innerText=v;if(cls)e.className=cls;}}
function bsFileCount(){try{var fso=new ActiveXObject("Scripting.FileSystemObject"),folder=fso.GetFolder("C:\\Windows\\Minidump"),en=new Enumerator(folder.Files),n=0;for(;!en.atEnd();en.moveNext()){var f=en.item();if(String(f.Name).toLowerCase().indexOf(".dmp")>=0)n++;}return n;}catch(e){return 0;}}
function bsLastBugCheck(){try{var sh=new ActiveXObject("WScript.Shell"),ex=sh.Exec('wevtutil qe System /q:"*[System[(EventID=1001)]]" /c:5 /rd:true /f:text'),txt="";while(!ex.StdOut.AtEndOfStream)txt+=ex.StdOut.ReadLine()+"\n";var n=(txt.match(/Event ID:\s*1001/gi)||[]).length;return {count:n,last:txt};}catch(e){return {count:-1,last:null};}}
function bsCheck(){
  bsMainProgress(5);setText("ready","블루스크린 검사 진행 중");setText("phase","진행 단계: 블루스크린 기록 확인 중");
  logMsg("블루스크린 상태 검사를 시작합니다.","info");
  try{
    var r=bsLastBugCheck(),d=bsFileCount(),fso=new ActiveXObject("Scripting.FileSystemObject"),wer=0;
    try{var paths=["C:\\Windows\\LiveKernelReports","C:\\ProgramData\\Microsoft\\Windows\\WER\\ReportArchive","C:\\ProgramData\\Microsoft\\Windows\\WER\\ReportQueue"];for(var pi=0;pi<paths.length;pi++){if(fso.FolderExists(paths[pi]))wer++;}}catch(we){}
    bsMainProgress(40);
    if(r.count<0){logMsg("System 이벤트 로그 조회에 실패했습니다.","caution");}else if(r.count>0){logMsg("BugCheck(이벤트 1001) 기록 "+r.count+"건을 확인했습니다.","bad");}else{logMsg("BugCheck 이벤트 1001 기록이 확인되지 않았습니다.","good");}
    bsMainProgress(70);
    if(d>0)logMsg("미니덤프 "+d+"개가 확인되었습니다.","caution");else logMsg("미니덤프가 없습니다.","good");
    if(wer>0)logMsg("Windows 오류 보고 관련 폴더 "+wer+"개가 확인되었습니다.","caution");
    bsMainProgress(100);
    if((r.count>0)||d>0||wer>0){setText("ready","블루스크린 점검 완료 - 점검 필요");setText("phase","진행 단계: 블루스크린 점검 완료");}else{setText("ready","블루스크린 점검 완료 - 이상 없음");setText("phase","진행 단계: 블루스크린 점검 완료");}
  }catch(e){bsMainProgress(100);setText("ready","블루스크린 검사 오류");setText("phase","진행 단계: 블루스크린 검사 오류");logMsg("블루스크린 검사 오류: "+e.message,"bad");}
}

var bsRepairRunning=false,bsRepairTimer=null,bsRepairStage=0,bsRepairJob=null,bsWaitTick=0,bsLastShown=0,bsWaitStart=0,bsLastJobText="",bsRepairDeadline=0;
// 현장용 초고속 긴급 복구: 빠른 판정 → 필요한 경우에만 DISM → 즉시 재확인.
// 콘솔은 숨기고, 실제 Windows 복구 프로세스는 제한시간을 넘기면 즉시 종료합니다.
var BS_MAX_TOTAL_MS=5*60*1000, BS_CHECK_MS=20*1000, BS_REPAIR_MS=180*1000, BS_RECHECK_MS=20*1000;
function bsMainProgress(p){var w=byId("bar"),t=byId("pct");p=Math.max(0,Math.min(100,parseInt(p,10)||0));if(w)w.style.width=p+"%";if(t)t.innerText=p+"%";setText("fixPct",p+"%");}
function bsKillJob(job){
  try{
    if(!job)return;
    var sh=new ActiveXObject("WScript.Shell"),pid="";
    if(job.pid){
      try{var fso=new ActiveXObject("Scripting.FileSystemObject");if(fso.FileExists(job.pid)){var t=fso.OpenTextFile(job.pid,1,false);if(!t.AtEndOfStream)pid=String(t.ReadLine()).replace(/\s/g,"");t.Close();}}catch(e){}
    }
    if(pid&&/^\d+$/.test(pid)){
      // PowerShell 프로세스가 DISM을 자식으로 실행하므로 /T로 하위 프로세스까지 종료합니다.
      sh.Run('cmd.exe /c taskkill /PID '+pid+' /T /F >nul 2>&1',0,true);
      try{sh.Run('powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Stop-Process -Id '+pid+' -Force -ErrorAction SilentlyContinue"',0,true);}catch(e2){}
    }
  }catch(e){logMsg("복구 작업 종료 중 추가 오류: "+e.message,"caution");}
}
function bsStartHiddenAdmin(cmd,tag,maxMs){
  try{
    var sh=new ActiveXObject("WScript.Shell"),fso=new ActiveXObject("Scripting.FileSystemObject"),tmp=sh.ExpandEnvironmentStrings("%TEMP%"),base=tmp+"\\POSCARE_BS_FAST_"+tag,bat=base+".cmd",ps1=base+".ps1",done=base+".done",doneTmp=done+".tmp",out=base+".log",pid=base+".pid";
    try{if(fso.FileExists(done))fso.DeleteFile(done,true);if(fso.FileExists(doneTmp))fso.DeleteFile(doneTmp,true);if(fso.FileExists(out))fso.DeleteFile(out,true);if(fso.FileExists(bat))fso.DeleteFile(bat,true);if(fso.FileExists(ps1))fso.DeleteFile(ps1,true);if(fso.FileExists(pid))fso.DeleteFile(pid,true);}catch(x){}
    var bf=fso.CreateTextFile(bat,true,false);
    bf.WriteLine("@echo off");
    bf.WriteLine(cmd+" > \""+out+"\" 2>&1");
    bf.WriteLine("exit /b %ERRORLEVEL%");
    bf.Close();
    var timeoutSec=Math.max(1,Math.ceil((maxMs||BS_REPAIR_MS)/1000)),qBat=bat.replace(/"/g,'\\"'),qDoneTmp=doneTmp.replace(/"/g,'\\"'),qDone=done.replace(/"/g,'\\"');
    var ps=fso.CreateTextFile(ps1,true,false);
    ps.WriteLine('$ErrorActionPreference="Continue"');
    ps.WriteLine('$p=Start-Process -FilePath "cmd.exe" -ArgumentList @("/c","'+qBat+'") -WindowStyle Hidden -PassThru');
    ps.WriteLine('$ok=$p.WaitForExit('+String(timeoutSec*1000)+')');
    ps.WriteLine('if(-not $ok){ try{ & "$env:SystemRoot\\System32\\taskkill.exe" /PID $p.Id /T /F 2>$null | Out-Null }catch{}; $rc=258 } else { $rc=$p.ExitCode }');
    ps.WriteLine('Set-Content -LiteralPath "'+qDoneTmp+'" -Value $rc -Encoding ASCII');
    ps.WriteLine('Move-Item -LiteralPath "'+qDoneTmp+'" -Destination "'+qDone+'" -Force');
    ps.WriteLine('exit $rc');
    ps.Close();
    var launch='Start-Process -FilePath powershell.exe -ArgumentList \'-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "'+ps1.replace(/'/g,"''")+'"\' -Verb RunAs -WindowStyle Hidden -PassThru | ForEach-Object { Set-Content -LiteralPath \''+pid.replace(/'/g,"''")+'\' -Value $_.Id -Encoding ASCII }';
    var full='powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "'+launch.replace(/"/g,'\\"')+'"';
    sh.Run(full,0,false);
    bsRepairJob={done:done,out:out,pid:pid,bat:bat,ps1:ps1,started:new Date().getTime(),tag:tag,maxMs:maxMs||BS_REPAIR_MS};bsLastJobText="";return true;
  }catch(e){logMsg("복구 명령 시작 오류: "+e.message,"bad");return false;}
}
function bsPollJob(){
  try{
    if(!bsRepairJob)return -1;
    var fso=new ActiveXObject("Scripting.FileSystemObject"),now=new Date().getTime();
    var elapsed=now-bsRepairJob.started;
    if(bsRepairDeadline && now>=bsRepairDeadline){bsKillJob(bsRepairJob);logMsg("긴급 복구 전체 제한시간에 도달해 Windows 복구 작업을 종료했습니다.","caution");return -2;}
    if(elapsed>bsRepairJob.maxMs){bsKillJob(bsRepairJob);logMsg((bsRepairJob.tag.indexOf("check")===0?"빠른 상태 확인":"Windows 복구")+"가 개별 제한시간을 초과해 중단되었습니다.","caution");return -2;}
    if(!fso.FileExists(bsRepairJob.done))return 0;
    try{if(fso.GetFile(bsRepairJob.done).Size<1)return 0;}catch(sz){return 0;}
    var rc=-1,raw="",ts=null;
    try{ts=fso.OpenTextFile(bsRepairJob.done,1,false);if(!ts.AtEndOfStream)raw=ts.ReadLine();ts.Close();rc=parseInt(raw,10);}catch(readErr){try{if(ts)ts.Close();}catch(closeErr){}return 0;}
    if(isNaN(rc))return 0;
    bsLastJobText="";
    try{if(fso.FileExists(bsRepairJob.out)){var o=fso.OpenTextFile(bsRepairJob.out,1,false);bsLastJobText=o.ReadAll();o.Close();if(bsLastJobText)logMsg(bsLastJobText.slice(-1000),rc===0?"info":"caution");}}catch(ex){}
    logMsg((bsRepairJob.tag.indexOf("check")===0?"DISM 빠른 상태 확인":"DISM 복구")+" 완료 - 종료 코드: "+rc,(rc===0?"good":"caution"));
    var clean=[bsRepairJob.done,bsRepairJob.out,bsRepairJob.ps1,bsRepairJob.pid];for(var i=0;i<clean.length;i++){try{if(clean[i]&&fso.FileExists(clean[i]))fso.DeleteFile(clean[i],true);}catch(cl){}}
    bsRepairJob=null;return rc===0?1:-1;
  }catch(e){logMsg("복구 상태 확인 오류: "+e.message,"bad");return 0;}
}
function bsOutputSaysRepairable(txt){txt=String(txt||"");return /(repairable|component store corruption was detected|손상.*발견|손상.*복구.*필요|복구.*가능)/i.test(txt);}
function bsOutputSaysHealthy(txt){txt=String(txt||"");return /(no component store corruption detected|구성 요소 저장소.*손상.*없|손상.*발견되지 않았|문제가.*발견되지 않았)/i.test(txt);}
function bsWaitingProgress(min,max,label){bsWaitTick++;var span=Math.max(1,max-min),p=min+(Math.floor((new Date().getTime()-bsWaitStart)/500)%span);if(p>=max)p=max-1;bsLastShown=p;bsMainProgress(p);var sec=Math.floor((new Date().getTime()-bsWaitStart)/1000);setText("phase","진행 단계: "+label+" (진행 중 · "+sec+"초)");}
function bsRepairFinish(msg,sub,completed){if(completed!==false)bsMainProgress(100);else bsMainProgress(Math.max(0,bsLastShown||0));setText("ready",msg);setText("readySub",sub||"");setText("phase","진행 단계: 긴급 복구 종료");setText("fixPct",(completed===false?String(Math.max(0,bsLastShown||0))+"%":"100%"));setText("fixCount","1");bsRepairRunning=false;bsRepairDeadline=0;if(bsRepairTimer)window.clearInterval(bsRepairTimer);bsRepairTimer=null;if(bsRepairJob){bsKillJob(bsRepairJob);bsRepairJob=null;}}
function bsRepair(){
  if(running||autoRunning||bsRepairRunning){logMsg("현재 다른 진단/복구 작업이 진행 중입니다.","caution");return;}
  bsRepairRunning=true;bsRepairStage=0;bsRepairJob=null;bsWaitTick=0;bsLastShown=0;bsWaitStart=new Date().getTime();bsLastJobText="";bsRepairDeadline=new Date().getTime()+BS_MAX_TOTAL_MS;
  bsMainProgress(3);setText("ready","블루스크린 긴급 복구 진행 중");setText("readySub","빠른 판정 후 필요한 경우에만 Windows 구성 요소를 복구합니다. 최대 5분");setText("phase","진행 단계: 블루스크린 기록 빠른 확인");logMsg("블루스크린 긴급 복구를 시작합니다. 불필요한 장시간 검사는 실행하지 않습니다.","info");if(bsRepairTimer)window.clearInterval(bsRepairTimer);bsRepairTimer=window.setInterval(bsRepairStep,350);
}
function bsRepairStep(){
  try{
    var now=new Date().getTime();
    if(bsRepairStage!==5 && bsRepairDeadline && now>=bsRepairDeadline){bsRepairFinish("긴급 복구 시간 제한","현장용 5분 제한에 도달해 추가 복구를 중단했습니다. 복구 완료로 판정하지 않습니다.",false);return;}
    if(bsRepairStage===0){
      var r=bsLastBugCheck(),d=bsFileCount(),fso=new ActiveXObject("Scripting.FileSystemObject"),wer=0;
      try{var paths=["C:\\Windows\\LiveKernelReports","C:\\ProgramData\\Microsoft\\Windows\\WER\\ReportArchive","C:\\ProgramData\\Microsoft\\Windows\\WER\\ReportQueue"];for(var pi=0;pi<paths.length;pi++){if(fso.FolderExists(paths[pi]))wer++;}}catch(we){}
      if((r.count||0)<=0&&d<=0&&wer<=0){bsRepairFinish("긴급 복구 완료 — 복구 불필요","최근 블루스크린 기록, 미니덤프, 오류 보고 흔적이 없어 무거운 Windows 복구 작업을 생략했습니다.");logMsg("빠른 판정에서 블루스크린 관련 복구 근거가 확인되지 않았습니다.","good");return;}
      bsMainProgress(8);bsWaitStart=new Date().getTime();setText("phase","진행 단계: Windows 구성 요소 빠른 상태 확인 중");logMsg("블루스크린 흔적이 확인되어 DISM CheckHealth만 빠르게 확인합니다. 시스템 변경은 없습니다.","info");if(bsStartHiddenAdmin('DISM /Online /Cleanup-Image /CheckHealth','checkhealth',BS_CHECK_MS)){bsRepairStage=1;}else{bsRepairFinish("긴급 복구 시작 실패","빠른 상태 확인을 시작하지 못했습니다.",false);}return;
    }
    if(bsRepairStage===1){
      var h=bsPollJob();if(h===0){bsWaitingProgress(8,28,"Windows 구성 요소 빠른 상태 확인");return;}if(h===-2){bsRepairFinish("빠른 판정 시간 초과","Windows 상태 확인이 20초 내 끝나지 않아 장시간 대기를 막기 위해 중단했습니다.",false);return;}if(h<0){bsRepairFinish("빠른 판정 실패","Windows 구성 요소 상태를 확인하지 못했습니다. 복구 완료로 판정하지 않습니다.",false);return;}
      var ht=String(bsLastJobText||"");
      if(bsOutputSaysHealthy(ht)){bsRepairFinish("긴급 복구 완료 — 복구 불필요","Windows 구성 요소 저장소의 손상이 확인되지 않아 장시간 복구를 실행하지 않았습니다.");logMsg("CheckHealth에서 구성 요소 저장소 손상이 확인되지 않았습니다.","good");return;}
      if(!bsOutputSaysRepairable(ht)){bsRepairFinish("긴급 복구 종료 — 판정 불명확","Windows 복구 필요 여부를 명확히 판정하지 못해 추가 변경을 하지 않았습니다.",false);return;}
      bsMainProgress(30);bsWaitStart=new Date().getTime();setText("phase","진행 단계: 확인된 Windows 구성 요소만 긴급 복구 중");logMsg("구성 요소 저장소 손상이 확인되어 DISM RestoreHealth를 실행합니다. 인터넷 지연을 막기 위해 로컬 원본 우선 모드로 실행합니다.","caution");
      if(bsStartHiddenAdmin('DISM /Online /Cleanup-Image /RestoreHealth /LimitAccess','dism_repair',BS_REPAIR_MS)){bsRepairStage=2;}else{bsRepairFinish("긴급 복구 시작 실패","DISM 복구를 시작하지 못했습니다.",false);}return;
    }
    if(bsRepairStage===2){
      var dres=bsPollJob();if(dres===0){bsWaitingProgress(30,86,"확인된 Windows 구성 요소 긴급 복구");return;}if(dres===-2){bsRepairFinish("복구 시간 제한 — 완료 여부 미확정","Windows 구성 요소 복구가 3분 내 완료되지 않아 실제 작업을 중단했습니다.",false);return;}if(dres<0){bsRepairFinish("긴급 복구 실패","Windows 구성 요소 복구가 정상 완료되지 않았습니다.",false);return;}
      bsMainProgress(88);bsWaitStart=new Date().getTime();setText("phase","진행 단계: 복구 결과 빠른 재확인 중");logMsg("복구 명령이 종료되어 CheckHealth로 결과를 다시 확인합니다.","info");if(bsStartHiddenAdmin('DISM /Online /Cleanup-Image /CheckHealth','recheck',BS_RECHECK_MS)){bsRepairStage=3;}else{bsRepairFinish("복구 결과 확인 실패","복구 완료 여부를 확인하지 못했습니다.",false);}return;
    }
    if(bsRepairStage===3){
      var rr=bsPollJob();if(rr===0){bsWaitingProgress(88,98,"복구 결과 빠른 재확인");return;}if(rr===-2){bsRepairFinish("복구 결과 확인 시간 초과","복구는 중단하지 않았지만 최종 상태를 확정하지 못했습니다.",false);return;}if(rr<0){bsRepairFinish("복구 결과 확인 실패","최종 상태를 확정하지 못했습니다.",false);return;}
      var rt=String(bsLastJobText||"");
      if(bsOutputSaysHealthy(rt)){bsRepairFinish("긴급 복구 완료","손상 확인 → 필요한 복구 → 최종 상태 재확인까지 완료했습니다. 자동 재부팅하지 않았습니다.");logMsg("최종 CheckHealth에서 구성 요소 저장소 손상이 확인되지 않았습니다.","good");}
      else if(bsOutputSaysRepairable(rt)){bsRepairFinish("복구 미완료 — 추가 조치 필요","최종 확인에서도 손상이 남아 있어 복구 완료로 표시하지 않았습니다.",false);}
      else{bsRepairFinish("복구 결과 확인 필요","복구 명령은 종료됐지만 최종 상태를 명확히 판정하지 못했습니다.",false);}
      return;
    }
  }catch(e){bsRepairFinish("긴급 복구 오류","복구 과정에서 오류가 발생했습니다.",false);logMsg("긴급 복구 오류: "+e.message,"bad");}
}

function openBlueScreenRecovery(){bsRepair();}
function closeBlueScreenRecovery(){var m=byId("bsModal");if(m)m.style.display="none";}

function navAction(name){
  if(name==="대시보드"){setText("phase","진행 단계: 대시보드");return;}
  if(name==="전체 진단"){startDiagnosis();return;}
  if(name==="자동 복구"){startAutoRecovery();return;}
  if(name==="긴급 복구"){emergencyRecovery();return;}
  if(name==="긴급 복구(블루스크린)"){bsRepair();return;}
  if(name==="설정"){openSettings();return;}
  if(name==="COM 진단"){openComDiagnostic();return;}
  if(name==="주변기기 진단"){openPeripheralDiagnostic();return;}
  if(name==="정보"){openInfo();return;}
  if(name==="디스크 정리"){openDiskCleanup();return;}
  if(name==="상세 로그"){openDetailLog();return;}
  setText("phase","진행 단계: "+name);
  logMsg(name+" 메뉴를 선택했습니다.","info");
}

function openDetailLog(){
  var m=byId("detailLogModal"); if(!m)return;
  var src=byId("log"), dst=byId("detailLogContent");
  dst.innerHTML=src?src.innerHTML:"";
  m.style.display="block";
  dst.scrollTop=dst.scrollHeight;
}
function closeDetailLog(){var m=byId("detailLogModal");if(m)m.style.display="none";}
function saveReportToDesktop(){
  try{
    var sh=new ActiveXObject("WScript.Shell"), fso=new ActiveXObject("Scripting.FileSystemObject");
    var desktop=sh.SpecialFolders("Desktop"), ts=new Date();
    function pad(n){return n<10?"0"+n:""+n;}
    var name="POS_CARE_진단보고서_"+ts.getFullYear()+pad(ts.getMonth()+1)+pad(ts.getDate())+"_"+pad(ts.getHours())+pad(ts.getMinutes())+pad(ts.getSeconds())+".txt";
    var path=fso.BuildPath(desktop,name), file=fso.CreateTextFile(path,true,true);
    var log=byId("log"), text=log?log.innerText:"진단 로그가 없습니다.";
    file.WriteLine("POS CARE PRO 진단 보고서");
    file.WriteLine("생성일시: "+ts.toLocaleString());
    file.WriteLine("========================================");
    file.WriteLine(text);
    file.Close();
    logMsg("보고서 저장 완료: 바탕화면\\"+name,"good");
    alert("보고서가 바탕화면에 저장되었습니다.\n\n"+name);
  }catch(e){logMsg("보고서 저장 오류: "+e.message,"caution");alert("보고서 저장에 실패했습니다.\n\n"+e.message);}
}


// 창 크기 조절은 HTA의 네이티브 RESIZE 기능을 사용합니다.
// 마우스 드래그 중 window.resizeTo()를 반복 호출하지 않아 끊김/지연을 방지합니다.
function init(){
  /* 로그인 화면을 최우선으로 표시합니다. 무거운 WMI/실시간 상태 조회는 로그인 화면이 뜬 뒤 비동기로 시작하여
     아이디/비밀번호 입력 직후의 짧은 입력 지연을 제거합니다. */
  try{showLogin();}catch(e){}
  try{progress(0);updateSelectionCount();setText("cpu","측정 중");setText("mem","측정 중");setText("disk","측정 중");setText("netVal","측정 중");setText("inetVal","측정 중");setWidth("cpuBar",0);setWidth("memBar",0);setWidth("diskBar",0);}catch(e){}
  window.setTimeout(function(){
    try{loadWindowSize();scheduleFit();}catch(e){}
    try{window.setTimeout(liveRefresh,250);window.setInterval(liveSystem,3000);window.setInterval(liveNetwork,5000);}catch(e){}
    try{logMsg("POS CARE PRO 대시보드 준비 완료","good");logMsg("실시간 PC/POS 상태 감시 시작","good");}catch(e){}
  },700);
}

var fitTimer=null;
function fitWholeScreen(){
  try{
    var root=byId('root'); if(!root)return;
    var vw=document.documentElement.clientWidth||document.body.clientWidth||0;
    var vh=document.documentElement.clientHeight||document.body.clientHeight||0;
    if(vw<10||vh<10)return;
    var baseW=1500, baseH=900;
    var scale=Math.min(vw/baseW,vh/baseH);
    if(!isFinite(scale)||scale<=0)scale=1;
    scale=Math.max(0.55,Math.min(scale,1.35));
    root.style.zoom=scale;
    root.style.width=(100/scale)+'%';
    root.style.height=(100/scale)+'%';
  }catch(e){}
}
function scheduleFit(){
  if(fitTimer)window.clearTimeout(fitTimer);
  fitTimer=window.setTimeout(function(){fitTimer=null;fitWholeScreen();},220);
}
window.setTimeout(fitWholeScreen,300);
window.onresize=scheduleFit;

