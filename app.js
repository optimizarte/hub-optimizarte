javascript:(async function(){'use strict';var HOST='_OAv7';var ex=document.getElementById(HOST);if(ex){ex.remove();return;}
var CTX=null;
try {
    var ctxEl=document.getElementById('_GAAN_ClientContext_');
    if(ctxEl) { CTX=JSON.parse(ctxEl.value); }
    else {
        var pageR=await fetch('/FE/ISC.Gaan.Actividad.FE.GridActividad/?_CIA=SCO');
        if(pageR.ok) {
            var html=await pageR.text();
            var mm=html.match(/name="_GAAN_ClientContext_"[^>]*value='([^']+)'/);
            if(!mm) mm=html.match(/name="_GAAN_ClientContext_"[^>]*value="([^"]+)"/);
            if(mm) {
                var cs=mm[1].replace(/&quot;/g, '"').replace(/&amp;/g, '&');
                CTX=JSON.parse(cs);
            }
        }
    }
} catch(e) {}
if(!CTX) { alert('No se pudo establecer la sesi\u00f3n con el CRM. Abre la pesta\u00f1a de Actividades primero.'); return; }
var USERS={'M441819E':{name:'Dany Hernandez',ini:'DH',color:'#DC0028',bg:'rgba(220,0,40,.18)'},'M354046Y':{name:'Fadoua Khachach',ini:'FK',color:'#3B82F6',bg:'rgba(59,130,246,.18)'},'MA48168T':{name:'Silvia Famoso',ini:'SF',color:'#10B981',bg:'rgba(16,185,129,.18)'}};var UIDS=['M441819E','M354046Y','MA48168T'];var TYPES={'Llamada':{i:'📞',c:'#3B82F6'},'Cita':{i:'📅',c:'#DC0028'},'Tarea':{i:'✅',c:'#DC0028'},'Carta':{i:'✉️',c:'#6366F1'},'Fax':{i:'📠',c:'#A855F7'},'Visita':{i:'🤝',c:'#10B981'},'Correo':{i:'📧',c:'#0EA5E9'},'Propensión':{i:'🎯',c:'#EC4899'},'WhatsApp':{i:'💬',c:'#25D366'},'Oportunidad':{i:'💼',c:'#8B5CF6'}};var DNAMES=['Lu','Ma','Mi','Ju','Vi','Sa','Do'];var MONTHS=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];var VTOS_KEY='oa_vtos_v1';function loadVtos(){try{return JSON.parse(localStorage.getItem(VTOS_KEY)||'[]');}catch(e){return[];}}
function saveVtos(items){try{localStorage.setItem(VTOS_KEY,JSON.stringify(items));}catch(e){}}
var ASG_KEY='oa_asg_v1';function loadAsg(){try{return JSON.parse(localStorage.getItem(ASG_KEY)||'{}');}catch(e){return{};}}function saveAsg(map){try{localStorage.setItem(ASG_KEY,JSON.stringify(map));}catch(e){}}

// ── OneDrive Overrides ─────────────────────────────────────
var OD=(function(){
  var TENANT='609cbec4-a6b7-422c-95c1-057586ed0026';
  var CLIENT='edf4bc1b-5274-4c3a-91fe-99a34d67fd12';
  var REDIRECT='https://grupoaplicaciones.gco.global/FE/ISC.Gaan.CRM.FE.Inicio/';
  var SCOPE='https://graph.microsoft.com/Files.ReadWrite offline_access User.Read';
  var FOLDER_PATH='/me/drive/root:/OPTIMIZARTE';
  var FILE_NAME='opticrm-overrides.json';
  var _folderID=null;
  var LS_TOK='od_tok_v1', LS_EXP='od_exp_v1';

  // Carregar token persistent de localStorage
  var _tok=null, _exp=0;
  try{
    var _lt=localStorage.getItem(LS_TOK);
    var _le=parseInt(localStorage.getItem(LS_EXP)||'0');
    if(_lt&&_le&&Date.now()<_le){_tok=_lt;_exp=_le;console.log('OD: token restaurat de localStorage');}
  }catch(e){}

  function _saveTok(tok,exp){
    _tok=tok;_exp=exp;
    try{localStorage.setItem(LS_TOK,tok);localStorage.setItem(LS_EXP,String(exp));}catch(e){}
  }
  function _clearTok(){
    _tok=null;_exp=0;
    try{localStorage.removeItem(LS_TOK);localStorage.removeItem(LS_EXP);}catch(e){}
  }

  function _authUrl(prompt){
    return 'https://login.microsoftonline.com/'+TENANT+'/oauth2/v2.0/authorize?'+
      new URLSearchParams({client_id:CLIENT,response_type:'token',redirect_uri:REDIRECT,
        scope:SCOPE,response_mode:'fragment',prompt:prompt||'none'}).toString();
  }
  async function _silentToken(){
    if(_tok&&Date.now()<_exp)return _tok;
    // Intent silenciós via iframe
    return new Promise(function(res){
      var f=document.createElement('iframe');
      f.style.cssText='position:fixed;width:0;height:0;border:0';
      f.src=_authUrl('none');
      var t=setTimeout(function(){try{document.body.removeChild(f);}catch(e){}res(null);},6000);
      f.onload=function(){
        try{
          var h=f.contentWindow.location.hash.substring(1);
          var p=new URLSearchParams(h);
          var tok=p.get('access_token');
          if(tok){var exp=Date.now()+(parseInt(p.get('expires_in')||'3600')-60)*1000;_saveTok(tok,exp);}
          clearTimeout(t);try{document.body.removeChild(f);}catch(e){}res(tok||null);
        }catch(e){clearTimeout(t);try{document.body.removeChild(f);}catch(e2){}res(null);}
      };
      document.body.appendChild(f);
    });
  }
  async function loginPopup(){
    var popup=window.open(_authUrl('select_account'),'ODLogin','width=500,height=600,left=400,top=100');
    return new Promise(function(res){
      var poll=setInterval(function(){
        try{
          var h=popup.location.hash.substring(1);
          if(h){var p=new URLSearchParams(h);var tok=p.get('access_token');
            if(tok){
              var exp=Date.now()+(parseInt(p.get('expires_in')||'3600')-60)*1000;
              _saveTok(tok,exp);
              clearInterval(poll);popup.close();res(true);
            }
          }
        }catch(e){}
        if(popup.closed){clearInterval(poll);res(!!_tok);}
      },300);
    });
  }
  async function _ensureFolder(tok){
    if(_folderID) return _folderID;
    try{
      var r=await fetch('https://graph.microsoft.com/v1.0'+FOLDER_PATH,{headers:{Authorization:'Bearer '+tok}});
      if(r.status===404){
        var rc=await fetch('https://graph.microsoft.com/v1.0/me/drive/root/children',{
          method:'POST',headers:{Authorization:'Bearer '+tok,'Content-Type':'application/json'},
          body:JSON.stringify({name:'OPTIMIZARTE',folder:{},['@microsoft.graph.conflictBehavior']:'rename'})
        });
        var jc=await rc.json();_folderID=jc.id;
        console.log('OD: carpeta creada ID:',_folderID);
      } else {var j=await r.json();_folderID=j.id;}
    }catch(e){console.log('OD ensureFolder error:',e.message);}
    return _folderID;
  }
  async function load(){
    var tok=_tok&&Date.now()<_exp?_tok:await _silentToken();
    if(!tok){console.log('OD: no token');return {};}
    try{
      var fid=await _ensureFolder(tok);if(!fid)return {};
      var r=await fetch('https://graph.microsoft.com/v1.0/me/drive/items/'+fid+':/'+FILE_NAME,{headers:{Authorization:'Bearer '+tok}});
      if(r.status===401){_clearTok();return {};}
      if(r.status===404){return {};}
      if(!r.ok){return {};}
      var meta=await r.json();
      var dl=meta['@microsoft.graph.downloadUrl'];if(!dl)return {};
      var r2=await fetch(dl);
      var d=await r2.json();
      console.log('OD: '+Object.keys(d).length+' overrides carregats');
      return d;
    }catch(e){console.log('OD load error:',e.message);return {};}
  }
  async function save(data){
    var tok=_tok&&Date.now()<_exp?_tok:await _silentToken();
    if(!tok){console.warn('OD: no token per guardar');return false;}
    try{
      var fid=await _ensureFolder(tok);if(!fid){console.warn('OD: no folder ID');return false;}
      var r=await fetch('https://graph.microsoft.com/v1.0/me/drive/items/'+fid+':/'+FILE_NAME+':/content',{
        method:'PUT',headers:{Authorization:'Bearer '+tok,'Content-Type':'text/plain'},
        body:JSON.stringify(data,null,2)
      });
      if(r.status===401){_clearTok();}
      if(r.ok)console.log('OD: '+Object.keys(data).length+' overrides guardats');
      else console.warn('OD save status:',r.status);
      return r.ok;
    }catch(e){console.log('OD save error:',e.message);return false;}
  }
  function isReady(){return !!_tok&&Date.now()<_exp;}
  return{load,save,loginPopup,isReady};
})();

// Overrides en memòria (carregats de OneDrive a l'inici)
var _overrides={};
async function _loadOverrides(){_overrides=await OD.load();window._overrides=_overrides;}
async function _saveOverride(id,dt){_overrides[id]=dt;window._overrides=_overrides;await OD.save(_overrides);}
function _applyOverrides(){S.all.forEach(function(a){var id=String(a._isOpo?a.IDOPOACT:a.IDACTIV);if(_overrides[id]){a.TIM_INICI=_overrides[id];}});}
var S={date:new Date(),view:'day',all:[],fil:[],aU:new Set(UIDS),aT:new Set(Object.keys(TYPES)),aS:new Set(['ab','ca','co']),sf:null,sq:'',rts:new Set(['act','opo']),drag:null,foc:{uid:'M441819E',filter:'all',types:new Set(),allDays:false},vtos:{items:loadVtos(),show:true},listFil:null,listRt:null,listTp:null};var pad=function(n){return String(n).padStart(2,'0');};var fmtD=function(d){return pad(d.getDate())+'/'+pad(d.getMonth()+1)+'/'+d.getFullYear();};var fmtISO=function(d){return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate());};var isToday=function(d){return fmtISO(d)===fmtISO(new Date());};var esc=function(s){return(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');};var gT=function(a){if(a._isOpo)return'Oportunidad';var d=(a.INDACTI_DESC||'').trim();var subj=(a.DESASUN||'').trim();if(d.includes('Llamada'))return'Llamada';if(d.includes('Cita'))return'Cita';if(d.includes('Tarea'))return'Tarea';if(d.includes('Carta'))return'Carta';if(d.includes('Fax'))return'Fax';if(d.includes('Visita'))return'Visita';if(d.includes('Correo'))return'Correo';if(d.includes('Propens'))return'Propensión';if(d.includes('WhatsApp')||d.includes('Whats'))return'WhatsApp';if(d.includes('Oportunidad')||d.includes('Opor'))return'Oportunidad';if(subj.match(/^!{0,2}[A-Z]?:?\s*[Ss]:?\s*OPO/)||subj.match(/^!!?[IDdSs]:?\s*OPO/)||subj.match(/OPO\s+[A-Z]/))return'Oportunidad';return'Tarea';};var gS=function(a){var d=(a.INDSITU_DESC||'').trim();return d.includes('Completad')?'co':d.includes('Cancelad')?'ca':'ab';};var gSL=function(s){return{ab:'Abierta',co:'Completada',ca:'Cancelada'}[s];};var gP=function(a){var d=(a.INDPRIO_DESC||'').trim();return d.includes('Alta')?{l:'Alta',c:'alta'}:(d.includes('Media')||d.includes('Normal'))?{l:'Med',c:'media'}:{l:'Baja',c:'baja'};};var gWk=function(d){var s=new Date(d);s.setDate(s.getDate()-((s.getDay()+6)%7));return Array.from({length:7},function(_,i){var r=new Date(s);r.setDate(s.getDate()+i);return r;});};var gMo=function(d){var y=d.getFullYear(),m=d.getMonth(),f=new Date(y,m,1),l=new Date(y,m+1,0),sw=(f.getDay()+6)%7,days=[];for(var i=sw-1;i>=0;i--){var x=new Date(y,m,-i);x._o=true;days.push(x);}for(var i=1;i<=l.getDate();i++)days.push(new Date(y,m,i));while(days.length%7!==0){var x=new Date(y,m+1,days.length-sw-l.getDate()+1);x._o=true;days.push(x);}return days;};var host=document.createElement('div');host.id=HOST;host.style.cssText='position:fixed;inset:0;z-index:2147483647;';document.body.appendChild(host);var SR=host.attachShadow({mode:'open'});var css=`:host{pointer-events:auto}*{box-sizing:border-box;margin:0;padding:0}#app{position:fixed;inset:0;display:grid;grid-template-columns:248px 1fr;grid-template-rows:min-content 1fr;font-family:Inter,system-ui,sans-serif;font-size:13px}#app[data-t="1"]{background:#0C0C10;color:#F2F2F8;--bg:#0C0C10;--bg2:#141419;--bg3:#1C1C24;--bg4:#23232E;--br:#2C2C38;--brl:#38384A;--t1:#F2F2F8;--t2:#B0B0CC;--t3:#8080A0;--t4:#4A4A62}#app[data-t="0"]{background:#F0F2F8;color:#141420;--bg:#F0F2F8;--bg2:#FFF;--bg3:#E8EAF4;--bg4:#DDE0EE;--br:#D0D4E8;--brl:#B8BEDD;--t1:#141420;--t2:#4A4A70;--t3:#7878A0;--t4:#A0A0C4}#hdr{grid-column:1/-1;min-height:52px;background:var(--bg2);border-bottom:1px solid var(--br);display:flex;align-items:center;padding:6px 14px;gap:10px;flex-wrap:wrap}.lm{width:28px;height:28px;background:#DC0028;border-radius:7px;display:grid;place-items:center;font-weight:800;font-size:13px;color:#fff;flex-shrink:0}.lt{font-weight:700;font-size:13px;color:var(--t1);line-height:1.2}.ls{font-size:9px;color:var(--t3);letter-spacing:.5px;text-transform:uppercase}.hd{width:1px;height:22px;background:var(--br);flex-shrink:0}.dn{display:flex;align-items:center;gap:6px}.dl{font-weight:600;font-size:13px;white-space:nowrap;min-width:180px;text-align:center;color:var(--t1);flex-shrink:0}.hr{margin-left:auto;display:flex;align-items:center;gap:4px;flex-shrink:0;flex-wrap:wrap}.cn{font-size:10px;color:var(--t3);white-space:nowrap}.vtabs{display:flex;gap:2px;background:var(--bg);padding:3px;border-radius:8px;flex-wrap:wrap}
button{font-family:inherit;cursor:pointer;border:none;background:transparent;font-size:13px;color:inherit;line-height:1}.dnb,.thb,.xb{display:grid;place-items:center}.dnb{width:26px;height:26px;border-radius:6px;border:1px solid var(--br);color:var(--t2);font-size:14px}.dnb:hover{background:var(--bg4);color:var(--t1)}.hoy{padding:4px 10px;border-radius:6px;border:1px solid var(--br);color:var(--t2);font-size:11px}.hoy:hover{background:var(--bg4);color:var(--t1)}.vt{padding:3px 10px;border-radius:5px;border:1px solid transparent;color:var(--t3);font-size:10.5px}.vt.on{background:var(--bg2);color:var(--t1);box-shadow:0 1px 4px rgba(0,0,0,.2)}.thb,.xb{width:28px;height:28px;border-radius:7px;border:1px solid var(--br);color:var(--t2);font-size:14px}.rfb{padding:4px 9px;border-radius:7px;border:1px solid var(--br);color:var(--t2);font-size:10.5px}.rtf{display:flex;gap:4px;}.rtfb{flex:1;padding:5px 0;border-radius:7px;border:1px solid var(--br);background:var(--bg3);color:var(--t3);font-size:10px;font-weight:600;cursor:pointer;text-align:center;font-family:inherit;transition:all.12s}.rtfb.on{background:rgba(16,185,129,.15);color:#10B981;border-color:rgba(16,185,129,.5);font-weight:700}.rtfb.opo.on{background:rgba(16,185,129,.15);color:#10B981;border-color:rgba(16,185,129,.5);font-weight:700}.rtfb.act.on{background:rgba(16,185,129,.15);color:#10B981;border-color:rgba(16,185,129,.4)}.srch{display:flex;align-items:center;gap:6px;flex-shrink:0}.srch-wrap{position:relative;display:flex;align-items:center}.srchi{display:block!important;visibility:visible!important;background:var(--bg3);border:1px solid var(--br);border-radius:7px;color:var(--t1);font-size:12px;padding:4px 28px 4px 10px;font-family:inherit;outline:none;width:14vw;min-width:110px;max-width:160px;transition:border-color.15s}.srch-clr{position:absolute;right:6px;background:none;border:none;cursor:pointer;color:var(--t3);font-size:13px;padding:0;line-height:1;opacity:0;pointer-events:none;transition:opacity.15s}.srch-clr.vis{opacity:1;pointer-events:auto}.srch-clr:hover{color:#EF4444}
.spill.on.co{background:rgba(16,185,129,.15);color:#10B981;border-color:rgba(16,185,129,.4)}.spill.on.ca{background:rgba(128,128,160,.15);color:#8080A0;border-color:rgba(128,128,160,.4)}.spill:not(.on){opacity:.55}.srchst:focus{border-color:var(--brl)}.thb:hover,.xb:hover,.rfb:hover{background:var(--bg4);color:var(--t1)}#sb{background:var(--bg2);border-right:1px solid var(--br);overflow-y:auto;padding:12px;display:flex;flex-direction:column;gap:0}#sb::-webkit-scrollbar{width:3px}#sb::-webkit-scrollbar-thumb{background:var(--br);border-radius:2px}.sbsec{margin-bottom:14px}.stit{font-size:9px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--t2);margin-bottom:8px;display:block}.vhdr{display:flex;align-items:center;justify-content:space-between;padding:8px 10px;border-radius:8px;background:rgba(220,0,40,.05);border:1px solid rgba(220,0,40,.12);margin-bottom:8px}.vtit2{display:flex;align-items:center;gap:6px;font-size:9px;font-weight:700;letter-spacing:.8px;text-transform:uppercase;color:#8080A0}.vcnt{background:rgba(86,86,106,.2);color:var(--t2);font-size:9px;font-weight:700;padding:1px 6px;border-radius:10px;min-width:18px;text-align:center}.vacts-hdr{display:flex;gap:4px;align-items:center}.vbtn{width:22px;height:22px;border-radius:5px;border:1px solid var(--br);background:transparent;color:var(--t3);font-size:13px;display:grid;place-items:center}.vbtn:hover{background:rgba(86,86,106,.2)}.vbtn.on{background:rgba(220,0,40,.12);border-color:#DC0028}.vprev{display:flex;flex-direction:column;gap:3px;margin-bottom:2px}.vpitem{border-left:3px solid var(--brl);background:var(--bg);border-radius:5px;padding:5px 8px;cursor:pointer}.vpitem:hover{background:var(--bg3)}.vptit{font-size:11px;font-weight:600;color:var(--t1);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.vpcli{font-size:10px;color:var(--t3);margin-top:1px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.mc{background:var(--bg);border-radius:7px;padding:10px}.mch{display:flex;align-items:center;justify-content:space-between;margin-bottom:8px}.mct{font-weight:600;font-size:10.5px;text-transform:capitalize;color:var(--t1);border-collapse:collapse;width:100%}.mct th{font-size:9px;font-weight:700;color:var(--t3);text-align:center;padding:3px 0}.mct td{text-align:center;padding:3px 1px;cursor:pointer;border-radius:4px;font-size:10.5px}.mct td:hover{background:var(--bg3)}.mct td.td{font-weight:800;color:#DC0028}.mct td.ts{background:#DC0028;color:#fff;border-radius:4px}.mc-nav{display:flex;align-items:center;justify-content:space-between;margin-bottom:6px}.mc-title{font-weight:700;font-size:12px;color:var(--t1)}.mc-arr{background:none;border:none;cursor:pointer;color:var(--t2);font-size:16px;padding:0 4px;line-height:1}.mc-arr:hover{color:var(--t1)}.mcn{display:flex;gap:2px}.mcn button{width:19px;height:19px;border-radius:4px;color:var(--t3);font-size:11px;display:grid;place-items:center}.mcn button:hover{color:var(--t1);background:var(--bg3)}.mcg{display:grid;grid-template-columns:repeat(7,1fr);gap:1px}.mcdh{text-align:center;font-size:8px;font-weight:700;color:var(--t3);padding:2px 0;text-transform:uppercase}.mcd{aspect-ratio:1;display:flex;align-items:center;justify-content:center;border-radius:4px;font-size:9.5px;cursor:pointer;color:var(--t2);position:relative}.mcd:hover{background:var(--bg4);color:var(--t1)}.mcd.tod{background:#DC0028;color:#fff;font-weight:700}.mcd.sel{background:rgba(220,0,40,.18);color:#DC0028;font-weight:600}.mcd.oth{color:var(--t4)}.mcd.has::after{content:'';width:3px;height:3px;border-radius:50%;background:#DC0028;position:absolute;bottom:1px;left:50%;transform:translateX(-50%)}.ur{display:flex;align-items:center;gap:8px;padding:7px 9px;border-radius:7px;cursor:pointer;border:1px solid transparent}.ur:hover{background:var(--bg3)}.ur.on{background:var(--bg);border-color:var(--br)}.av{display:grid;place-items:center;font-weight:700;flex-shrink:0;border-radius:7px}.uinfo{flex:1;min-width:0}.un{font-size:11px;font-weight:600;color:var(--t1);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.uc{font-size:9px;color:var(--t3);margin-top:1px}.utog{width:16px;height:16px;border-radius:4px;border:1.5px solid var(--br);display:grid;place-items:center;font-size:9px;color:transparent;flex-shrink:0;margin-left:auto}.utog.on{background:#DC0028;border-color:#DC0028;color:#fff}.tfrow{display:flex;align-items:center;gap:6px;padding:4px 8px;border-radius:5px;cursor:pointer}.tfrow:hover{background:var(--bg3)}.tdot{width:6px;height:6px;border-radius:50%;flex-shrink:0}.tfsep{height:1px;background:var(--br);margin:6px 4px 4px;}.select.dsel{background:var(--bg3)!important;color:var(--t1)!important;border:1px solid var(--br);border-radius:5px;font-size:11px;padding:1px 3px}select.dsel option{background:#1C1C24!important;color:#F2F2F8!important;}.tl{font-size:11px;color:var(--t2);flex:1}.tc{font-size:9px;color:var(--t3);background:var(--bg);padding:1px 5px;border-radius:9px}.sfrow{display:flex;align-items:center;gap:6px;padding:4px 8px;border-radius:5px;cursor:pointer}.sfrow:hover{background:var(--bg3)}.togbtn{border-radius:5px;border:1px solid var(--br);color:var(--t3);font-size:9.5px;padding:2px 8px}.togbtn:hover{color:var(--t1);border-color:var(--brl)}#mn{display:flex;flex-direction:column;overflow:hidden;background:var(--bg)}#sbar{display:flex;align-items:center;gap:12px;padding:7px 14px;background:var(--bg2);border-bottom:1px solid var(--br);overflow-x:auto;flex-shrink:0}#sbar::-webkit-scrollbar{height:0}.sest{display:flex;align-items:center;gap:2px;flex-shrink:0}.sfrow2{display:flex;align-items:center;gap:5px;padding:3px 8px;border-radius:6px;cursor:pointer;border:1px solid transparent;white-space:nowrap}.sfrow2:hover{background:var(--bg3);border-color:var(--br)}.sfrow2.soff{opacity:.4}.tdot2{width:4px;height:4px;border-radius:50%;flex-shrink:0;background:#DC0028;display:inline-block;vertical-align:middle}.sel2{font-size:10px;color:var(--t2)}.stc2{font-size:11px;font-weight:700;color:var(--t2);min-width:16px;text-align:right}.si{display:flex;align-items:center;gap:5px;white-space:nowrap;cursor:pointer;padding:3px 7px;border-radius:6px;border:1px solid transparent;flex-shrink:0}.si:hover{background:var(--bg3);border-color:var(--br)}.si.on{background:var(--bg3);border-color:var(--brl)}.sil{font-size:10px;color:var(--t2)}.siv{font-size:13px;font-weight:700}.sdv{width:1px;height:13px;background:var(--br);flex-shrink:0}#cnt{flex:1;min-height:0;overflow:hidden;display:flex;flex-direction:column}.dcols{display:grid;flex:1;overflow:hidden;border-top:1px solid var(--br)}.ucol{border-right:1px solid var(--br);display:flex;flex-direction:column;overflow:hidden;min-height:0}.ucol:last-child{border-right:none}.uch{padding:10px 13px;border-bottom:1px solid var(--br);display:flex;align-items:center;gap:8px;background:var(--bg2);flex-shrink:0}.ucb{flex:1;overflow-y:auto;padding:8px;display:flex;flex-direction:column;gap:5px}.ucb::-webkit-scrollbar{width:3px}.ucb::-webkit-scrollbar-thumb{background:var(--br);border-radius:2px}.ucb.dov{background:rgba(220,0,40,.06);outline:2px dashed rgba(220,0,40,.4);outline-offset:-4px;border-radius:8px}.ucb.vdov{background:rgba(86,86,106,.08);outline:2px dashed rgba(220,0,40,.5);outline-offset:-4px;border-radius:8px}.act{background:var(--bg2);border:1px solid var(--br);border-radius:7px;padding:8px 10px 8px 13px;cursor:grab;position:relative;overflow:visible;animation:fu.18s ease both;transition:border-color.12s,background.12s,transform.12s}.act:active{cursor:grabbing}.act::before{content:'';position:absolute;left:0;top:0;bottom:0;width:3px}.act:hover{border-color:var(--brl);background:var(--bg3);transform:translateX(2px)}.act.dn{opacity:.4}.act.pp{border-style:dashed}.act.drg{opacity:.3;transform:rotate(1deg)}@keyframes fu{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:translateY(0)}}.atop{display:flex;align-items:flex-start;gap:6px;margin-bottom:4px}.atime{font-size:10px;font-weight:600;color:var(--t2);white-space:nowrap;flex-shrink:0;min-width:42px}.asubj{font-size:12px;font-weight:600;color:var(--t1);flex:1;line-height:1.4;word-break:break-word}.ab{flex-shrink:0;font-size:8px;font-weight:700;padding:2px 5px;border-radius:3px;text-transform:uppercase;letter-spacing:.3px}.ab.ab{background:rgba(220,0,40,.12);color:#FF7A95}.ab.co{background:rgba(16,185,129,.12);color:#10B981}.ab.ca{background:rgba(86,86,106,.15);color:var(--t3)}.acli{font-size:10px;color:var(--t2);margin-top:1px}.aobs{font-size:9.5px;color:var(--t2);margin-top:4px;line-height:1.5;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}.afoot{display:flex;align-items:center;justify-content:space-between;margin-top:6px}.atype{display:flex;align-items:center;gap:4px;font-size:10px;color:var(--t2)}.apr{font-size:7.5px;font-weight:700;padding:1px 4px;border-radius:3px;text-transform:uppercase;margin-left:3px}.apr.alta{background:rgba(239,68,68,.14);color:#EF4444}.apr.media{background:rgba(245,158,11,.14);color:#DC0028}.apr.baja{background:rgba(86,86,106,.14);color:var(--t3)}.aact{display:flex;gap:3px;opacity:0;transition:opacity.13s}.act:hover.aact{opacity:1}.abtn{width:20px;height:20px;border-radius:4px;border:1px solid var(--br);background:var(--bg);color:var(--t3);font-size:10px;cursor:pointer;display:grid;place-items:center;line-height:1}.abtn:hover{color:var(--t1);border-color:var(--brl)}.abtn.ok:hover{color:#10B981;border-color:#10B981}.abtn.go:hover{color:#DC0028;border-color:#DC0028}.adel{width:22px;height:22px;border-radius:4px;border:1px solid rgba(239,68,68,.4);background:rgba(239,68,68,.06);color:#EF4444;font-size:12px;cursor:pointer;display:grid;place-items:center;line-height:1}.adel:hover{color:#EF4444;border-color:#EF4444;background:rgba(239,68,68,.08)}.wdel{width:20px;height:20px;border-radius:4px;border:1px solid rgba(239,68,68,.4);background:rgba(239,68,68,.1);color:#EF4444;font-size:14px;line-height:1;cursor:pointer;display:flex;align-items:center;justify-content:center;flex-shrink:0}.wact:hover.wdel{opacity:1}.wdel:hover{border-color:#EF4444;background:rgba(239,68,68,.25)}.ldel{width:22px;height:22px;border-radius:4px;border:1px solid rgba(239,68,68,.4);background:rgba(239,68,68,.06);color:#EF4444;font-size:12px;cursor:pointer;display:grid;place-items:center;flex-shrink:0}.ldel:hover{color:#EF4444;border-color:#EF4444;background:rgba(239,68,68,.08)}.vact{background:var(--bg2);border:1px solid rgba(245,158,11,.25);border-left:3px solid#DC0028;border-radius:7px;padding:8px 10px 8px 11px;cursor:grab;animation:fu.15s ease both;transition:border-color.12s}.vact:hover{border-color:var(--brl);background:var(--bg3)}.vact.drg{opacity:.3;transform:rotate(1deg)}.vtit{font-size:12px;font-weight:600;color:var(--t1);margin-bottom:3px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.vcli2{font-size:10px;color:var(--t2);margin-bottom:5px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.vfoot{display:flex;align-items:center;justify-content:space-between}.vdate{font-size:9px;color:var(--t3)}.vdel{width:18px;height:18px;border-radius:3px;border:1px solid var(--br);background:transparent;color:var(--t4);font-size:10px;cursor:pointer;display:grid;place-items:center;opacity:0;transition:opacity.1s}.vact:hover.vdel{opacity:1}.vdel:hover{color:#EF4444;border-color:#EF4444}.empty{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;color:var(--t3);gap:7px;padding:18px;text-align:center}.empi{font-size:24px;opacity:.3}.empt{font-size:10.5px;line-height:1.7;color:var(--t3)}.ld{display:flex;align-items:center;justify-content:center;gap:8px;padding:28px;color:var(--t3);font-size:12px;flex:1}.sp{width:15px;height:15px;border:2px solid var(--br);border-top-color:#DC0028;border-radius:50%;animation:spin.7s linear infinite}@keyframes spin{to{transform:rotate(360deg)}}
.wkwrap{flex:1;min-height:0;display:flex;flex-direction:column;overflow:hidden}.foc-wrap{display:flex;height:100%;flex-direction:column;overflow:hidden}.foc-bar{display:flex;align-items:center;gap:8px;padding:8px 14px;background:var(--bg2);border-bottom:1px solid var(--br);flex-shrink:0;flex-wrap:wrap}.foc-label{font-size:10px;font-weight:700;color:var(--t3);text-transform:uppercase;letter-spacing:.5px;margin-right:4px}.foc-usr{display:flex;gap:4px}.foc-upill{padding:4px 10px;border-radius:20px;border:2px solid transparent;font-size:11px;font-weight:700;cursor:pointer;font-family:inherit;transition:all.12s;background:var(--bg3);color:var(--t3)}.foc-upill.on{color:#fff}.foc-sep{width:1px;background:var(--br);margin:0 4px;align-self:stretch}.foc-type{display:flex;gap:3px;flex-wrap:wrap}.foc-tpill{padding:3px 9px;border-radius:12px;border:1px solid var(--br);font-size:10px;font-weight:600;cursor:pointer;font-family:inherit;background:var(--bg3);color:var(--t3);transition:all.12s}.foc-tpill.on{background:rgba(220,0,40,.12);color:#FF7A95;border-color:rgba(220,0,40,.3)}.foc-tpill.on-opo{background:rgba(139,92,246,.12);color:#8B5CF6;border-color:rgba(139,92,246,.3)}.foc-list{flex:1;overflow-y:auto;padding:10px 14px;display:grid;grid-template-columns:repeat(4,1fr);grid-auto-rows:min-content;gap:8px;align-content:start}.foc-card{background:var(--bg2);border:1px solid var(--br);border-radius:8px;padding:10px 12px 10px 14px;cursor:pointer;border-left:3px solid var(--br);transition:border-color.12s,background.12s;position:relative;min-width:0;display:flex;flex-direction:column;gap:3px}.foc-card:hover{background:var(--bg3);border-color:var(--brl)}.foc-card.sdone{opacity:0.3;filter:grayscale(70%)}.foc-card.foc-opo{border-left-color:#8B5CF6}.foc-top{display:flex;align-items:flex-start;gap:6px;flex-wrap:wrap}.foc-time{font-size:10px;font-weight:700;color:var(--t2);white-space:nowrap;flex-shrink:0;min-width:48px}.foc-subj{font-size:13px;font-weight:700;color:var(--t1);flex:1;line-height:1.3;word-break:break-word}.foc-st{font-size:9px;font-weight:700;padding:2px 7px;border-radius:8px;flex-shrink:0}.foc-st.ab{background:rgba(16,185,129,.12);color:#10B981}.foc-st.co{background:rgba(59,130,246,.12);color:#3B82F6}.foc-st.ca{background:rgba(239,68,68,.12);color:#EF4444}.foc-cli{font-size:11px;color:var(--t2);margin-top:2px}.foc-meta{display:flex;align-items:center;gap:6px;margin-top:5px}.foc-badge{font-size:9px;font-weight:600;padding:1px 6px;border-radius:4px;background:var(--bg3);color:var(--t3)}@media(min-width:1400px){.foc-list{grid-template-columns:repeat(5,1fr)}}@media(max-width:900px){.foc-list{grid-template-columns:repeat(2,1fr)}}.foc-empty{grid-column:1/-1;display:flex;align-items:center;justify-content:center;flex-direction:column;gap:6px;opacity:.4;padding:40px}.foc-empty-ico{font-size:32px}.foc-empty-txt{font-size:12px;color:var(--t3)}.tl-wrap{flex:1;min-height:0;overflow-y:auto;padding:12px 16px;display:flex;flex-direction:column;gap:12px}.tl-user{border-radius:10px;background:var(--bg2);border:1px solid var(--br);overflow:hidden}.tl-uhdr{display:flex;align-items:center;gap:8px;padding:8px 12px;background:var(--bg3);border-bottom:1px solid var(--br)}.tl-uname{font-size:12px;font-weight:700;color:var(--t1)}.tl-cnt{font-size:10px;color:var(--t3)}.tl-scroll{overflow-x:auto;overflow-y:hidden}.tl-track{display:flex;gap:0;min-height:80px;position:relative}.tl-day{flex:0 0 56px;border-right:1px solid var(--br);padding:4px 3px;display:flex;flex-direction:column;gap:2px}.tl-day.tl-today{background:rgba(220,0,40,.05)}.tl-dhd{font-size:9px;font-weight:700;color:var(--t3);text-align:center;margin-bottom:3px;padding-bottom:2px;border-bottom:1px solid var(--br)}.tl-dhd.td{color:#DC0028}.tl-item{border-radius:3px;padding:2px 4px;font-size:9px;font-weight:600;cursor:pointer;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;border-left:2px solid transparent;background:var(--bg3);color:var(--t2);line-height:1.4}.tl-item:hover{color:var(--t1);opacity:1}.tl-item.tl-opo{background:rgba(139,92,246,.1);border-left-color:#8B5CF6;color:#8B5CF6}.tl-noact{font-size:10px;color:var(--t4);text-align:center;padding:4px 0}.wkhdr{display:grid;grid-template-columns:repeat(7,1fr);scrollbar-gutter:stable;border-bottom:1px solid var(--br);flex-shrink:0;background:var(--bg2)}.wkdh{padding:10px 8px 8px;text-align:center;cursor:pointer;border-right:1px solid var(--br)}.wkdh:last-child{border-right:none}.wkdh:hover{background:var(--bg3)}.wkdh.wkt{background:rgba(220,0,40,.08)}.wkdh.wks{background:rgba(220,0,40,.14)}.wkdn{font-size:8.5px;font-weight:700;color:var(--t3);text-transform:uppercase;letter-spacing:.4px}.wkdd{font-size:18px;font-weight:700;color:var(--t1);margin:2px 0;line-height:1.2}.wkdh.wkt.wkdd{color:#DC0028}.wkdc{font-size:9px;color:var(--t3)}.wkgrid{display:grid;grid-template-columns:repeat(7,1fr);flex:1;min-height:0;overflow-y:scroll;scrollbar-gutter:stable}.wkcol{border-right:1px solid var(--br);display:flex;flex-direction:column;gap:4px;padding:6px;min-height:80px}.wkcol:last-child{border-right:none}.wkcol.dov{background:rgba(220,0,40,.06);outline:2px dashed rgba(220,0,40,.4);outline-offset:-3px;border-radius:5px}.wact{background:var(--bg2);border:1px solid var(--br);border-left:3px solid transparent;border-radius:5px;padding:5px 7px 5px 9px;cursor:grab;font-size:10.5px;animation:fu.15s ease both}.wact.sdone,.act-card.sdone{opacity:0.28;filter:grayscale(80%)}.wact:hover{border-color:var(--brl);background:var(--bg3)}.wact.dn{opacity:.35}.warow{display:flex;align-items:center;gap:4px}.wusel{font-size:8.5px;font-weight:700;border-radius:4px;border:1px solid var(--br);background:var(--bg3);color:var(--t2);padding:1px 3px;cursor:pointer;height:18px;max-width:34px;outline:none;font-family:inherit}.wusel:hover{border-color:var(--brl);color:var(--t1)}.lusel{font-size:10px;font-weight:600;border-radius:5px;border:1px solid var(--br);background:var(--bg3);color:var(--t2);padding:2px 6px;cursor:pointer;height:24px;outline:none;font-family:inherit}.lusel:hover{border-color:var(--brl);color:var(--t1)}.wat{font-size:9px;color:var(--t3);flex-shrink:0;min-width:32px}.was{flex:1;font-weight:600;color:var(--t1)}.wau{width:16px;height:16px;border-radius:3px;display:grid;place-items:center;font-size:7.5px;font-weight:700;flex-shrink:0}.mowrap{flex:1;min-height:0;display:flex;flex-direction:column;overflow:hidden}.mohdr{display:grid;grid-template-columns:repeat(7,1fr);background:var(--bg2);border-bottom:1px solid var(--br);flex-shrink:0}.modh{text-align:center;padding:7px 4px;font-size:9px;font-weight:700;color:var(--t2);text-transform:uppercase;letter-spacing:.5px}.mogrid{display:grid;grid-template-columns:repeat(7,1fr);grid-auto-rows:1fr;flex:1;min-height:0;overflow-y:auto}.moday{border-right:1px solid var(--br);border-bottom:1px solid var(--br);padding:5px 6px;display:flex;flex-direction:column;gap:2px;cursor:pointer;min-height:70px}.moday:hover{background:var(--bg3)}.moday.mt{background:rgba(220,0,40,.07)}.moday.ms{background:rgba(220,0,40,.12)}.moday.mo{opacity:.35}.moday.dov{outline:2px dashed rgba(220,0,40,.5);outline-offset:-3px;border-radius:4px;background:rgba(220,0,40,.06)}.modn{font-size:12px;font-weight:700;color:var(--t1);line-height:1.2}.moday.mt.modn{color:#DC0028}.modots{display:flex;flex-wrap:wrap;gap:2px;margin-top:3px;pointer-events:none}.modot{width:7px;height:7px;border-radius:50%;pointer-events:none}.molbl{font-size:9px;color:var(--t2);border-radius:3px;padding:1px 4px;background:var(--bg3);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;pointer-events:none}.lview{flex:1;min-height:0;overflow-y:auto;padding:12px 16px;display:flex;flex-direction:column;gap:4px}.lview::-webkit-scrollbar{width:3px}.lview::-webkit-scrollbar-thumb{background:var(--br);border-radius:2px}.ldg{font-size:9px;font-weight:700;color:var(--t2);text-transform:uppercase;letter-spacing:.8px;padding:9px 0 4px}.la{background:var(--bg2);border:1px solid var(--br);border-radius:7px;padding:9px 13px;cursor:pointer;display:grid;grid-template-columns:16px 52px 1fr auto 26px 70px 24px;align-items:center;gap:8px}.la:hover{background:var(--bg3);border-color:var(--brl)}.las{font-size:12px;font-weight:600;color:var(--t1);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.lcli{font-size:11px;color:var(--t2);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.vmodal{position:absolute;inset:0;background:rgba(0,0,0,.6);display:flex;align-items:center;justify-content:center;z-index:10}.det-modal{position:absolute;inset:0;background:rgba(0,0,0,.65);display:flex;align-items:center;justify-content:center;z-index:20}.det-box{background:var(--bg2);border:1px solid var(--brl);border-radius:14px;width:min(720px,95vw);max-height:88vh;display:flex;flex-direction:column;overflow:hidden}.det-hdr{padding:16px 20px 0;flex-shrink:0}.det-title{font-size:15px;font-weight:700;color:var(--t1);margin-bottom:4px;line-height:1.3}.det-sub{font-size:11px;color:var(--t3);display:flex;gap:10px;flex-wrap:wrap;margin-bottom:12px}.det-tabs{display:flex;gap:2px;border-bottom:1px solid var(--br);padding:0 20px;flex-shrink:0;background:var(--bg2)}.det-tab{padding:8px 14px;font-size:12px;font-weight:600;color:var(--t3);cursor:pointer;border-bottom:2px solid transparent;margin-bottom:-1px;white-space:nowrap}.det-tab.on{color:#DC0028;border-bottom-color:#DC0028}.det-tab:hover{color:var(--t1)}.det-body{flex:1;overflow-y:auto;padding:16px 20px}.det-body::-webkit-scrollbar{width:4px}.det-body::-webkit-scrollbar-thumb{background:var(--br);border-radius:2px}.det-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px 20px}.det-field{display:flex;flex-direction:column;gap:3px}.det-field.full{grid-column:1/-1}.det-label{font-size:9.5px;font-weight:700;color:var(--t3);text-transform:uppercase;letter-spacing:.5px}.det-value{font-size:12px;color:var(--t1);line-height:1.4;word-break:break-word}.det-value.empty{color:var(--t4);font-style:italic}.det-value a{color:#3B82F6;text-decoration:none}.det-value a:hover{text-decoration:underline}.det-obs{font-size:12px;color:var(--t1);line-height:1.6;white-space:pre-wrap;background:var(--bg3);border-radius:7px;padding:12px;border:1px solid var(--br)}.det-footer{display:flex;gap:8px;justify-content:flex-end;padding:12px 20px;border-top:1px solid var(--br);flex-shrink:0;background:var(--bg2)}.det-btn{padding:7px 14px;border-radius:7px;font-size:12px;font-weight:600;cursor:pointer;font-family:inherit;border:1px solid var(--br);color:var(--t2);background:var(--bg3)}.det-btn:hover{background:var(--bg4);color:var(--t1)}.det-btn.pri{background:#DC0028;color:#fff;border-color:#DC0028}.det-btn.pri:hover{background:#c00020}.det-btn.ok{background:rgba(16,185,129,.12);color:#10B981;border-color:rgba(16,185,129,.3)}.det-btn.ok:hover{background:rgba(16,185,129,.2)}.det-btn.del{background:rgba(239,68,68,.1);color:#EF4444;border-color:rgba(239,68,68,.3)}.det-btn.del:hover{background:rgba(239,68,68,.2)}.det-badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700;text-transform:uppercase}.det-badge.ab{background:rgba(220,0,40,.12);color:#FF7A95}.det-badge.co{background:rgba(16,185,129,.12);color:#10B981}.det-badge.ca{background:rgba(86,86,106,.15);color:var(--t3)}.vmbox{background:var(--bg2);border:1px solid var(--brl);border-radius:12px;padding:22px;width:320px}.vmbox h3{font-size:14px;font-weight:700;color:var(--t1);margin-bottom:3px}.vmbox p{font-size:11px;color:var(--t3);margin-bottom:14px;line-height:1.5}.vf{display:flex;flex-direction:column;gap:4px;margin-bottom:10px}.vf label{font-size:10px;font-weight:600;color:var(--t3);text-transform:uppercase;letter-spacing:.4px}.vf input,.vf select,.vf textarea{background:var(--bg3);border:1.5px solid var(--br);border-radius:6px;padding:7px 10px;color:var(--t1);font-size:13px;font-family:inherit;width:100%}.vf input:focus,.vf select:focus,.vf textarea:focus{outline:none;border-color:#DC0028}.vf textarea{resize:none;height:52px}.vmbtns{display:flex;gap:8px;justify-content:flex-end;margin-top:14px}.vmcancel{padding:6px 13px;border-radius:7px;border:1px solid var(--br);color:var(--t2);font-size:12px}.vmcancel:hover{background:var(--bg4);color:var(--t1)}.vmsave{padding:6px 15px;border-radius:7px;background:#DC0028;color:#000;font-weight:700;font-size:12px;border:none}.vmsave:hover{background:#FF1A3C}`;var styleEl=document.createElement('style');styleEl.textContent=css;SR.appendChild(styleEl);var app=document.createElement('div');app.id='app';app.setAttribute('data-t','0');var q=function(id){return SR.getElementById(id);};var qa=function(sel){return Array.from(SR.querySelectorAll(sel));};var SB=document.createElement('div');SB.id='sb';SB.innerHTML='<div class="sbsec">'+
'<div class="vhdr">'+
'<div class="vtit2">📋 <span>OPO/VTOS</span></div>'+
'<div class="vacts-hdr">'+
'<span class="vcnt" id="vcnt">0</span>'+
'<button class="vbtn" id="veye" onclick="_OAv.vtosToggle()" title="Ver columna en Vista Día">👁</button>'+
'<button class="vbtn" onclick="_OAv.vtosNew()" title="Nueva entrada">+</button>'+
'</div></div>'+
'<div id="vprev" class="vprev"></div>'+
'</div>'+
'<div class="sbsec"><span class="stit">Calendario</span><div class="mc" id="mc"></div></div>'+
'<div class="sbsec"><span class="stit">Colaboradores</span>'+
UIDS.map(function(uid){var u=USERS[uid];return('<div class="ur on" id="u'+uid+'" onclick="_OAv.tu(\''+uid+'\',this)">'+
'<div class="av" style="width:28px;height:28px;background:'+u.bg+';color:'+u.color+';font-size:9.5px">'+u.ini+'</div>'+
'<div class="uinfo"><div class="un">'+u.name+'</div><div class="uc" id="uc'+uid+'">0 act.</div></div>'+
'<div class="utog on" id="ut'+uid+'">✓</div></div>');}).join('')+
'</div>'+
'<div class="sbsec">'+
'<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">'+
'<span class="stit" style="margin-bottom:0">Tipo</span>'+
'<button class="togbtn" onclick="_OAv.ta()" id="ta">Desel. todo</button></div>'+
'<div id="tf"></div></div>'+
'';var H=document.createElement('div');H.id='hdr';H.innerHTML='<div class="lm">O</div>'+
'<div><div class="lt">OPTIMIZARTE</div><div class="ls">Agenda Comercial</div></div>'+
'<div class="hd"></div>'+
'<div class="dn"><button class="dnb" onclick="_OAv.nd(-1)">&#8249;</button><div class="dl" id="dl"></div><button class="dnb" onclick="_OAv.nd(1)">&#8250;</button></div>'+
'<button class="hoy" onclick="_OAv.gt()">Hoy</button>'+
'<div class="vtabs">'+
'<button class="vt on" onclick="_OAv.sv(\'day\',this)">Día</button>'+
'<button class="vt" onclick="_OAv.sv(\'week\',this)">Semana</button>'+
'<button class="vt" onclick="_OAv.sv(\'month\',this)">Mes</button>'+
'<button class="vt" onclick="_OAv.sv(\'list\',this)">Lista</button>'+
'<button class="vt" onclick="_OAv.sv(\'timeline\',this)">⏱ Línea</button>'+
'<button class="vt" onclick="_OAv.sv(\'focus\',this)">🎯 Focus</button>'+'<button class="vt" id="tab-kanban" onclick="_OAv.sv(\'kanban\',this)" style="margin-left:8px">📊 Embudo</button>'+
'</div>'+
'<div class="hr"><span class="cn" id="cn">Cargando...</span>'+
'<div class="srch">'+
'<input class="srchi" id="srchi" type="text" placeholder="🔍 Buscar..." oninput="_OAv.srch(this.value)" onkeydown="if(event.key===\'Escape\'){this.value=\'\';_OAv.srch(\'\')}">'+
'</div>'+
'<button class="rfb" onclick="_OAv.fa()">↺ Recargar</button>'+
'<button class="rfb" id="diagb" onclick="_OAv.diag()" style="background:rgba(245,158,11,.12);border-color:rgba(245,158,11,.35);color:#F59E0B">Diag</button>'+
'<button class="thb" onclick="_OAv.tg()" id="tb">☀️</button>'+
'<button class="xb" onclick="_OAv.close()">×</button></div>';var MN=document.createElement('div');MN.id='mn';var SBAR=document.createElement('div');SBAR.id='sbar';SBAR.innerHTML='<div class="si" onclick="_OAv.ssf(\'all\',this)"><span class="sil">Total</span><span class="siv" id="sv0" style="color:var(--t1)">—</span></div><div class="sdv"></div>'+
'<div class="si" onclick="_OAv.ssf(\'pen\',this)"><span class="sil">Pendientes</span><span class="siv" id="sv1" style="color:#FF7A95">—</span></div><div class="sdv"></div>'+
'<div class="si" onclick="_OAv.ssf(\'don\',this)"><span class="sil">Completadas</span><span class="siv" id="sv2" style="color:#10B981">—</span></div><div class="sdv"></div>'+
'<div class="si" onclick="_OAv.ssf(\'vis\',this)"><span class="sil">Visitas</span><span class="siv" id="sv3" style="color:#10B981">—</span></div><div class="sdv"></div>'+
'<div class="si" onclick="_OAv.ssf(\'cit\',this)"><span class="sil">Citas</span><span class="siv" id="sv4" style="color:#DC0028">—</span></div>'+
'<div style="flex:1"></div>'+
'<div class="sdv"></div>'+
'<div class="sest">'+
'<div class="sfrow2" id="sfab" style="cursor:default"><div class="tdot2" style="background:#FF7A95"></div><span class="sel2" style="color:var(--t1);font-weight:700">En Curso (Abiertas)</span><span class="stc2" id="scab">0</span></div>'+
'</div>';var CNT=document.createElement('div');CNT.id='cnt';CNT.innerHTML='<div class="ld"><div class="sp"></div><span>Cargando actividades...</span></div>';MN.appendChild(SBAR);MN.appendChild(CNT);app.appendChild(H);app.appendChild(SB);app.appendChild(MN);SR.appendChild(app);function mkDrag(el,id,isV){el.draggable=true;el.addEventListener('dragstart',function(e){S.drag={id:id,vtos:!!isV};el.classList.add('drg');e.dataTransfer.effectAllowed='move';e.dataTransfer.setData('text/plain',id);});el.addEventListener('dragend',function(){el.classList.remove('drg');qa('.dov,.vdov').forEach(function(x){x.classList.remove('dov','vdov');});S.drag=null;});}
function mkDrop(el,fn,vt){el.addEventListener('dragover',function(e){e.preventDefault();var c=vt?'vdov':'dov';if(!el.classList.contains(c))el.classList.add(c);});el.addEventListener('dragleave',function(e){if(!el.contains(e.relatedTarget)){el.classList.remove('dov','vdov');}});el.addEventListener('drop',function(e){e.preventDefault();el.classList.remove('dov','vdov');if(!S.drag)return;fn(S.drag);});}
var USER_FETCH_CONFIG=[{uid:'M441819E',selMed:'M441819E',codAsoc:null,targetUid:'M441819E'},{uid:'M354046Y',selMed:'M354046Y',codAsoc:null,targetUid:'M354046Y'},{uid:'MA48168T',selMed:'-1',codAsoc:'353621',targetUid:'MA48168T'}];async function fetchUser(uid){var cfg=null;for(var i=0;i<USER_FETCH_CONFIG.length;i++){if(USER_FETCH_CONFIG[i].uid===uid){cfg=USER_FETCH_CONFIG[i];break;}}
if(!cfg)cfg={uid:uid,selMed:uid,codAsoc:null,targetUid:uid};var msg={SelectorMediador:{ClaveSeleccionada:cfg.selMed},SelectorVista:{ClaveSeleccionada:'208'},Concepto:{Valor:null},CheckMisActividades:{Valor:false},AE_COD_MEDIADOR_RACF:'00017167',CIAGRUPO:'SCO',UID_RACF:uid,ENTIDAD:'Appointment',OPERATIVA:'GridActividades',NAVEGACION:'CARGA_INICIAL',COD_MEDIADOR_LOGADO:'-1',UsuarioRol:'Oficina'};if(cfg.codAsoc)msg.CodAsociados={ClaveSeleccionada:cfg.codAsoc};var ctx=Object.assign({},CTX,{init:'ISC.Gaan.Actividad.FE.GridActividad'});var body=new URLSearchParams({'_GAAN_Mensaje_Pantalla_':JSON.stringify(msg),'_GAAN_Comando_':'buscar','_GAAN_ClientContext_':JSON.stringify(ctx),'_GAAN_SPA_':'true'});try{var ac=new AbortController();setTimeout(function(){ac.abort();},8000);var r=await fetch('/FE/ISC.Gaan.Actividad.FE.GridActividad/',{method:'POST',signal:ac.signal,headers:{'Content-Type':'application/x-www-form-urlencoded'},body:body.toString()});if(!r.ok)return[];var json=await r.json();var data=(json.message&&json.message.ListaResultados&&json.message.ListaResultados.FuenteDatos)||[];return data.map(function(a){return Object.assign({},a,{_uid:cfg.targetUid});});}catch(e){console.warn('OA fetch',uid,e);return[];}}
async function crmPost(mod,cmd,msg){try{var ctx=Object.assign({},CTX,{init:mod});var body=new URLSearchParams({'_GAAN_Mensaje_Pantalla_':JSON.stringify(msg),'_GAAN_Comando_':cmd,'_GAAN_ClientContext_':JSON.stringify(ctx),'_GAAN_SPA_':'true'});var r=await fetch('/FE/'+mod+'/',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:body.toString()});return r.ok;}catch(e){return false;}}
async function reschedCRM(id,nd){
  var a=S.all.find(function(x){return x.IDACTIV===id;});if(!a)return;
  // Guardar override a OneDrive (persistent multi-usuari)
  _saveOverride(id,nd);
  // Intentar escriure al CRM (activitats)
  if(!a._isOpo){
    await crmPost('ISC.Gaan.Actividad.FE.DetalleActividad','guardar',{IDACTIV:id,TIM_INICI:nd,TIM_FI:nd,UID_RACF:a._uid||'M441819E',AE_COD_MEDIADOR_RACF:'00017167',CIAGRUPO:'SCO'});
  }
}
function applyF(){S.fil=S.all.filter(function(a){var uid=a._uid||a.UID_RACF||'M441819E';if(!S.aU.has(uid))return false;if(!S.aT.has(gT(a)))return false;if(S.sf==='pen')return gS(a)==='ab';if(S.sf==='don')return gS(a)==='co';if(S.sf==='vis')return gT(a)==='Visita'&&S.aS.has(gS(a));if(S.sf==='cit')return gT(a)==='Cita'&&S.aS.has(gS(a));if(!S.aS.has(gS(a)))return false;var isOpo=gT(a)==='Oportunidad';if(isOpo&&!S.rts.has('opo'))return false;if(!isOpo&&!S.rts.has('act'))return false;if(S.sq){var q=S.sq.toLowerCase();var hay=(a.DESASUN||'').toLowerCase()+(a.CLIENTE||'').toLowerCase()+(a.CODMEDI_DESC||'').toLowerCase();if(hay.indexOf(q)<0)return false;}
return true;});updC();render();}
function set(id,v){var e=q(id);if(e)e.textContent=v;}
function updC(){UIDS.forEach(function(uid){set('uc'+uid,S.all.filter(function(a){return(a._uid||a.UID_RACF||'M441819E')===uid;}).length+' act.');});set('scab',S.all.filter(function(a){return gS(a)==='ab';}).length);set('scco',0);set('scca',S.all.filter(function(a){return gS(a)==='ca';}).length);set('sv0',S.fil.length);set('sv1',S.fil.filter(function(a){return gS(a)==='ab';}).length);set('sv2',S.fil.filter(function(a){return gS(a)==='co';}).length);set('sv3',S.fil.filter(function(a){return gT(a)==='Visita';}).length);set('sv4',S.fil.filter(function(a){return gT(a)==='Cita';}).length);updTF();rMC();updVtos();}
function openCRM(mod,idKey,idVal){var f=document.createElement('form');f.method='POST';f.action='/FE/'+mod+'/?_CIA=SCO';f.target='_blank';var in1=document.createElement('input');in1.type='hidden';in1.name='_GAAN_Mensaje_Pantalla_';var msg={};msg[idKey]=idVal;in1.value=JSON.stringify(msg);var in2=document.createElement('input');in2.type='hidden';in2.name='_GAAN_ClientContext_';in2.value=JSON.stringify(Object.assign({},CTX,{init:mod}));f.appendChild(in1);f.appendChild(in2);document.body.appendChild(f);f.submit();f.remove();}
function _updHeader(){if(!S.all||!S.all.length){var cnEl=q('cn');if(cnEl)cnEl.textContent='Cargando...';return;}
var activeUids=Array.from(S.aU);var totalActs=S.all.filter(function(a){return!a._isOpo&&activeUids.indexOf(a._uid||a.UID_RACF)>=0;}).length||0;var totalOpos=S.all.filter(function(a){return a._isOpo&&activeUids.indexOf(a._uid||a.UID_RACF)>=0;}).length||0;var cnEl=q('cn');if(cnEl){cnEl.textContent='✓ '+totalActs+' act · '+totalOpos+' OPOs';}}
function updTF(){var cnts={};Object.keys(TYPES).forEach(function(t){cnts[t]=0;});S.all.forEach(function(a){var t=gT(a);if(t in cnts)cnts[t]++;});var opoKey='Oportunidad';var mkRow=function(k,e){return'<div class="tfrow" data-t="'+k+'" style="opacity:'+(S.aT.has(k)?1:.35)+'"><div class="tdot" style="background:'+e.c+'"></div><span class="tl">'+e.i+' '+k+'</span><span class="tc">'+(cnts[k]||0)+'</span></div>';};var opoEntry=TYPES[opoKey];var opoRow=opoEntry?mkRow(opoKey,opoEntry)+'<div class="tfsep"></div>':'';var rest=Object.entries(TYPES).filter(function(e){return e[0]!==opoKey;}).map(function(e){return mkRow(e[0],e[1]);}).join('');var tf=q('tf');if(!tf)return;tf.innerHTML=opoRow+rest;tf.onclick=function(ev){var row=ev.target.closest('.tfrow');if(row)_OAv.tt(row.getAttribute('data-t'),row);};var allOn=Object.keys(TYPES).every(function(t){return S.aT.has(t);});var btn=q('ta');if(btn)btn.textContent=allOn?'Desel. todo':'Sel. todo';}
function updVtos(){var items=S.vtos.items;set('vcnt',items.length);var pv=q('vprev');if(!pv)return;if(items.length===0){pv.innerHTML='<div style="font-size:10px;color:var(--t4);padding:2px 8px">Sin entradas VTOS</div>';return;}
pv.innerHTML=items.slice(0,3).map(function(it){return'<div class="vpitem"><div class="vptit">'+esc(it.asunto||'—')+'</div><div class="vpcli">'+esc(it.cliente||'—')+'</div></div>';}).join('')+(items.length>3?'<div style="font-size:10px;color:var(--t3);padding:2px 8px">+'+(items.length-3)+' más</div>':'');}
var mcOff=0;function rMC(){var base=new Date(S.date);base.setMonth(base.getMonth()+mcOff);var yr=base.getFullYear(),mo=base.getMonth();var first=new Date(yr,mo,1),lastDay=new Date(yr,mo+1,0).getDate();var startDow=(first.getDay()+6)%7;var today=new Date();var mc=q('mc');if(!mc)return;var dots={};S.all.forEach(function(a){var t=(a.TIM_INICI||'').slice(0,10);if(t){var p=t.split('/');if(p.length===3&&parseInt(p[1])-1===mo&&parseInt(p[2])===yr)dots[parseInt(p[0])]=true;}});var rows='<tr>';['L','M','X','J','V','S','D'].forEach(function(d){rows+='<th>'+d+'</th>';});rows+='</tr><tr>';var col=0;for(var i=0;i<startDow;i++){rows+='<td></td>';col++;}
for(var d=1;d<=lastDay;d++){var isToday=today.getDate()===d&&today.getMonth()===mo&&today.getFullYear()===yr;var isSel=S.date.getDate()===d&&S.date.getMonth()===mo&&S.date.getFullYear()===yr;var hasDot=dots[d];rows+='<td class="'+(isToday?'td':'')+(isSel?' ts':'')+'" onclick="_OAv.jd('+yr+','+mo+','+d+')" style="cursor:pointer;text-align:center">'
+'<div style="display:flex;flex-direction:column;align-items:center;line-height:1.2">'+d+'<span class="tdot2" style="'+(hasDot?'':'visibility:hidden')+'margin-top:1px"></span></div>'+'</td>';col++;if(col%7===0&&d<lastDay)rows+='</tr><tr>';}
rows+='</tr>';mc.innerHTML='<div class="mc-nav">'+
'<button onclick="_OAv.mcM(-1)" class="mc-arr">‹</button>'+
'<span class="mc-title">'+MONTHS[mo]+' '+yr+'</span>'+
'<button onclick="_OAv.mcM(1)" class="mc-arr">›</button>'+
'</div>'+
'<table class="mct"><tbody>'+rows+'</tbody></table>';}
function rHdr(){var d=S.date;var txt=S.view==='week'?(function(){var wk=gWk(d);return wk[0].getDate()+' '+MONTHS[wk[0].getMonth()]+' – '+wk[6].getDate()+' '+MONTHS[wk[6].getMonth()]+' '+d.getFullYear();})():S.view==='month'?MONTHS[d.getMonth()]+' '+d.getFullYear():['domingo','lunes','martes','miércoles','jueves','viernes','sábado'][d.getDay()]+', '+d.getDate()+' '+MONTHS[d.getMonth()]+' '+d.getFullYear();var el=q('dl');if(el)el.textContent=txt;}
function card(a,idx){var tp=gT(a),ti=TYPES[tp]||TYPES['Tarea'],st=gS(a),pr=gP(a);var t1=a.TIM_INICI_HORA||'',t2=a.TIM_FI_HORA||'',pol=(a.REFEXTE_POL||'').trim(),obs=(a.OBSACT||'').trim();var isProp=tp==='Propensión',isDone=st==='co';var u_uid=a._uid||a.UID_RACF||'M441819E';var u=USERS[u_uid]||USERS['M441819E'];var el=document.createElement('div');el.className='act'+(isDone?' dn':'')+(isProp?' pp':'');el.style.animationDelay=(idx*.03)+'s';el.innerHTML='<div style="position:absolute;left:0;top:0;bottom:0;width:3px;background:'+ti.c+';border-radius:3px 0 0 3px"></div>'+
'<div class="atop"><span class="atime">'+t1+(t2&&t2!==t1?'–'+t2:'')+'</span><span class="asubj">'+esc(a.DESASUN||'Sin asunto')+'</span><span class="ab '+st+'">'+gSL(st)+'</span></div>'+
(a.CLIENTE?'<div class="acli">· '+esc(a.CLIENTE)+(pol?' <span style="font-size:8.5px;color:var(--t3)">· '+pol+'</span>':'')+'</div>':'')+
(obs?'<div class="aobs">'+esc(obs)+'</div>':'')+
'<div class="afoot"><div class="atype">'+ti.i+' '+tp+'<span class="apr '+pr.c+'">'+pr.l+'</span></div>'+
'<div class="aact" id="aact_'+a.IDACTIV+'"></div></div>';var aactDiv=el.querySelector('.aact')||el.querySelector('#aact_'+a.IDACTIV);if(aactDiv){if(!isDone&&!isProp){var bOk=document.createElement('button');bOk.className='abtn ok';bOk.textContent='✓';bOk.title='Completar';bOk.addEventListener('click',function(e){e.stopPropagation();_OAv.ca(a.IDACTIV);});aactDiv.appendChild(bOk);}
if(a.REFNUMPERS){var bCli=document.createElement('button');bCli.className='abtn';bCli.textContent='👤';bCli.title='Ver cliente';bCli.addEventListener('click',function(e){e.stopPropagation();_OAv.oc(a.REFNUMPERS);});aactDiv.appendChild(bCli);}
var bGo=document.createElement('button');bGo.className='abtn go';bGo.textContent='↗';bGo.title='Abrir en CRM';bGo.addEventListener('click',function(e){e.stopPropagation();_OAv.oa(a.IDACTIV);});aactDiv.appendChild(bGo);var bDel=document.createElement('button');bDel.className='adel';bDel.textContent='×';bDel.title='Cancelar actividad';bDel.addEventListener('click',function(e){e.stopPropagation();_OAv.da(a.IDACTIV);});aactDiv.appendChild(bDel);var dsel=document.createElement('select');dsel.className='dsel';var u_bg2=(USERS[u_uid]||USERS['M441819E']).bg;var u_col2=(USERS[u_uid]||USERS['M441819E']).color;dsel.style.cssText='font-size:8px;font-weight:700;border-radius:4px;border:1px solid '+u_col2+';background:'+u_bg2+';color:'+u_col2+';cursor:pointer;outline:none;appearance:none;-webkit-appearance:none;padding:1px 3px;max-width:32px;margin-left:2px';UIDS.forEach(function(uid2){var op=document.createElement('option');op.value=uid2;op.textContent=USERS[uid2].ini;if(uid2===u_uid)op.selected=true;dsel.appendChild(op);});dsel.addEventListener('mousedown',function(e){e.stopPropagation();});dsel.addEventListener('click',function(e){e.stopPropagation();});dsel.addEventListener('change',function(e){e.stopPropagation();_OAv.rua(a.IDACTIV,this.value);});aactDiv.appendChild(dsel);}
el.addEventListener('click',function(e){if(e.target.tagName==='SELECT'||e.target.classList.contains('abtn')||e.target.classList.contains('adel'))return;showDetail(a.IDACTIV);});mkDrag(el,a.IDACTIV,false);return el;}
function showDetail(id) {
    try {
        var a = typeof id === 'object' ? id : S.all.find(function(x) { return x.IDACTIV === id; });
        if (!a) return;
        var existing = document.getElementById('vmodal');
        if(existing) existing.remove();
        
        var ov = document.createElement('div');
        ov.className = 'vmodal';
        ov.id = 'vmodal';
        
        var box = document.createElement('div');
        box.className = 'vmbox shadow-lg';
        box.style.display = 'flex';
        box.style.flexDirection = 'column';
        box.style.width = 'min(850px, 95vw)';
        box.style.maxWidth = 'none';
        box.style.maxHeight = '90vh';
        
        var isOpo = a._isOpo || gT(a) === 'Oportunidad';
        var st = gS(a);
        var ti = gP(a);
        var uname = 'Desconocido';
        try { uname = USERS[a._uid||a.UID_RACF||'M441819E'].name; } catch(e){ uname = a.CODMEDI_DESC || 'Desconocido'; }
        
        var hd = document.createElement('div');
        hd.style.cssText = 'padding:16px 20px;border-bottom:1px solid var(--br);display:flex;justify-content:space-between;align-items:center;background:var(--bg2)';
        hd.innerHTML = '<div style="display:flex;align-items:center;gap:12px">' +
            '<div style="width:36px;height:36px;border-radius:8px;background:' + (isOpo ? 'rgba(139,92,246,0.1)' : 'var(--bg3)') + ';display:flex;align-items:center;justify-content:center;font-size:20px">' + ((TYPES[gT(a)] || TYPES['Tarea']).i || '') + '</div>' +
            '<div>' +
            '<div style="font-weight:700;font-size:16px;color:var(--t1)">' + (a.DESASUN || 'Detalle de Registro') + '</div>' +
            '<div style="font-size:13px;color:var(--t2);margin-top:2px">' + (isOpo ? 'Oportunidad de Venta' : 'Actividad ' + gT(a)) + ' • ' + (a.TIM_INICI || '') + '</div>' +
            '</div></div>' +
            '<button class="xb" id="det-close-' + a.IDACTIV + '" style="cursor:pointer">×</button>';
        box.appendChild(hd);
        
        var cnt = document.createElement('div');
        cnt.style.cssText = 'padding:20px;flex:1;overflow-y:auto;display:flex;flex-direction:column;gap:16px;background:var(--bg1)';
        
        var obsTxt = (a.OBSERVACIONES || a.DS_DESCRIPCION || a.DESCRIPCION || a.OBSACT || '').trim() || 'Sin detalles.';
        var htmlLinks = '';
        var rawLinks = a.DS_ENLACE || a.ENLACE || '';
        if (rawLinks) {
            var links = rawLinks.split('|');
            links.forEach(function(lk) {
                var pts = lk.split('::');
                if (pts.length >= 2) htmlLinks += '<a href="' + pts[1] + '" target="_blank" style="display:inline-block;padding:8px 12px;background:rgba(59,130,246,0.1);color:#3B82F6;border-radius:6px;text-decoration:none;font-weight:600;font-size:12px;margin-right:8px;margin-bottom:8px">📎 ' + pts[0] + '</a>';
            });
        }
        
        if (isOpo) {
            // ── CARD 1: Dades principals + dates ──────────────────────────
            var card1 = document.createElement('div');
            card1.style.cssText = 'background:var(--bg2);border:1px solid var(--br);border-radius:10px;padding:16px';

            var diesActius = '—';
            var cierreAlerta = false;
            var datopClean = (a.DATOPOR||'').replace(/[^0-9\/]/g,'').trim();
            if (datopClean) {
                var partsC = datopClean.split('/');
                if (partsC.length === 3 && partsC[2].length === 4) {
                    var dCreacio = new Date(parseInt(partsC[2]), parseInt(partsC[1])-1, parseInt(partsC[0]));
                    var avui2 = new Date(); avui2.setHours(0,0,0,0);
                    var da = Math.round(Math.abs(avui2 - dCreacio) / 86400000);
                    if (!isNaN(da)) diesActius = da;
                }
            }
            if (a.TIM_FI) {
                var partsF = (a.TIM_FI||'').replace(/[^0-9\/]/g,'').split('/');
                if (partsF.length === 3 && partsF[2].length === 4) {
                    var dCierre = new Date(parseInt(partsF[2]), parseInt(partsF[1])-1, parseInt(partsF[0]));
                    var avui3 = new Date(); avui3.setHours(0,0,0,0);
                    if (dCierre - avui3 < 60 * 24 * 3600000) cierreAlerta = true;
                }
            }

            var lbl = function(txt){ return '<div style="font-size:9.5px;color:var(--t3);text-transform:uppercase;font-weight:700;letter-spacing:.4px;margin-bottom:3px">'+txt+'</div>'; };
            var lbl = function(ico,txt){ return '<div style="display:flex;align-items:center;gap:5px;font-size:9.5px;color:var(--t3);text-transform:uppercase;font-weight:700;letter-spacing:.4px;margin-bottom:4px"><span>'+ico+'</span><span>'+txt+'</span></div>'; };
            var val = function(txt,bold,col){ return '<div style="font-size:'+(bold?'13':'12')+'px;color:'+(col||'var(--t1)')+(bold?';font-weight:600':'')+';padding-left:18px">'+(txt||'—')+'</div>'; };

            var selBaseStyle='font-size:10.5px;font-weight:600;border-radius:5px;border:1px solid var(--br);background:#f4f4f4;color:#333;padding:2px 4px;cursor:pointer;font-family:inherit;outline:none;margin-left:18px;box-sizing:border-box';
            var selKbStyle=selBaseStyle+';width:64%';
            var selMedStyle=selBaseStyle+';width:86%';

            // Kanban state
            var kbVal = localStorage.getItem('kb_'+a.IDOPOACT);
            if(!kbVal){ kbVal = (function(){if(datopClean){var pp2=datopClean.split('/');if(pp2.length===3){var ts=new Date(parseInt(pp2[2]),parseInt(pp2[1])-1,parseInt(pp2[0])).getTime();if(ts<new Date().setHours(0,0,0,0))return'espera';}}return'analisis';})(); }
            var kbOpts=[['espera','⏳ En Espera'],['analisis','🔍 En Anàlisi'],['tarifa','💰 Tarificació'],['presentar','📋 Presentar'],['respuesta','⏱ Pdte. Resp.'],['emitir','✉️ Per Emetre'],['cierre','🏁 Cierre']];
            var kbHtml='<select id="kb-sel-'+a.IDOPOACT+'" style="'+selKbStyle+'">';
            kbOpts.forEach(function(o){kbHtml+='<option value="'+o[0]+'"'+(kbVal===o[0]?' selected':'')+'>'+o[1]+'</option>';});
            kbHtml+='</select>';
            var stIcon = st==='co'?'✅':st==='ca'?'❌':'⏰';

            // Mediador dropdown
            var medSelHtml='<select id="med-sel-'+a.IDOPOACT+'" style="'+selMedStyle+'">';
            UIDS.forEach(function(uid){var u=USERS[uid];medSelHtml+='<option value="'+uid+'"'+(a._uid===uid?' selected':'')+'>'+u.ini+' '+u.name+'</option>';});
            medSelHtml+='</select>';

            var c1h = '<div style="display:grid;grid-template-columns:1.5fr 1fr 1fr 0.35fr;gap:12px;margin-bottom:12px;align-items:start">';
            c1h += '<div>'+lbl('👤','Client')+'<div style="font-size:13px;font-weight:600;color:var(--t1);padding-left:18px">'+esc(a.CLIENTE||'—')+'</div>'+(a.REFNUMPERS?'<div style="font-size:10px;color:var(--t3);padding-left:18px">'+a.REFNUMPERS+'</div>':'')+'</div>';
            c1h += '<div>'+lbl('📊','Kanban')+kbHtml+'</div>';
            c1h += '<div>'+lbl('🛡️','Mediador')+medSelHtml+'</div>';
            c1h += '<div style="text-align:center">'+lbl('📌','')+'<div style="font-size:24px;line-height:1;text-align:center">'+stIcon+'</div></div>';
            c1h += '</div>';

            var cierreColor = cierreAlerta ? '#DC0028' : 'var(--t1)';
            var diesColor = typeof diesActius==='number'&&diesActius>365?'#EF4444':typeof diesActius==='number'&&diesActius>180?'#F59E0B':'var(--t2)';
            var proxRaw=(a.TIM_INICI||a.TIMFREA||'');
            var proxISO='';
            if(proxRaw){var pp=proxRaw.split('/');if(pp.length===3)proxISO=pp[2]+'-'+(pp[1]+'').padStart(2,'0')+'-'+(pp[0]+'').padStart(2,'0');}
            var cierreRaw=(a.TIM_FI||'').replace(/[^0-9\/]/g,'');
            var cierreISO='';
            if(cierreRaw){var cp=cierreRaw.split('/');if(cp.length===3)cierreISO=cp[2]+'-'+(cp[1]+'').padStart(2,'0')+'-'+(cp[0]+'').padStart(2,'0');}
            var inputStyle='font-size:13px;font-weight:700;color:#1a1a2e;background:#f4f4f4;border:1.5px solid rgba(59,130,246,.35);border-radius:6px;padding:2px 4px;font-family:inherit;cursor:pointer;outline:none;width:68%;box-sizing:border-box;margin-left:18px;box-shadow:0 1px 4px rgba(59,130,246,.1)';
            c1h += '<div style="display:grid;grid-template-columns:1fr 1fr 0.7fr 1fr;gap:8px;padding-top:10px;border-top:1px solid var(--br)">';
            c1h += '<div>'+lbl('📅','Pròxima Acció')+'<input type="date" id="prox-inp-'+a.IDOPOACT+'" value="'+proxISO+'" style="'+inputStyle+';" /></div>';
            c1h += '<div>'+lbl('🎯','Cierre Estimado')+'<input type="date" id="cierre-inp-'+a.IDOPOACT+'" value="'+cierreISO+'" style="'+inputStyle+';'+(cierreAlerta?'color:#DC0028;border-color:rgba(220,0,40,.4)':'')+'" /></div>';
            c1h += '<div style="text-align:center">'+lbl('⏱️','Dies')+'<div style="font-size:14px;font-weight:700;color:'+diesColor+';text-align:center;padding-top:2px">'+diesActius+(typeof diesActius==='number'?'<span style="font-size:10px;color:var(--t3)">d</span>':'')+'</div></div>';
            c1h += '<div style="text-align:center">'+lbl('🗓️','Creació')+'<div style="font-size:12px;color:var(--t1);text-align:center;padding-top:2px">'+val(datopClean||a.DATOPOR||'—').replace('padding-left:18px','text-align:center;padding-left:0')+'</div></div>';
            c1h += '</div>';

            card1.innerHTML = c1h;
            cnt.appendChild(card1);

            // ── CARD V360: Visió 360 col·lapsable ─────────────────────────
            if (a.REFNUMPERS) {
                var refPts = a.REFNUMPERS.split('/');
                var v360np = refPts[0]||'';
                var v360rp = refPts[1]||'000';
                var v360BaseUrl = 'https://catalanaaplicaciones.gco.global/FE/PER.Clientes.Menu.FE.Busqueda/?_CIA=SCO&NUMPERS='+v360np+'&REFPERS='+v360rp+'&ORIGEN=PER';

                var cardV360 = document.createElement('div');
                cardV360.style.cssText = 'border:1px solid var(--br);border-radius:10px;overflow:hidden;background:var(--bg2)';

                var v360hdr = document.createElement('button');
                v360hdr.style.cssText = 'width:100%;padding:11px 16px;display:flex;align-items:center;justify-content:space-between;background:rgba(139,92,246,.05);border:none;cursor:pointer;font-family:inherit';
                v360hdr.innerHTML = '<div style="display:flex;align-items:center;gap:8px"><span style="font-size:14px">🔍</span><div style="text-align:left"><div style="font-size:11px;font-weight:700;color:var(--t1)">Visió 360 · '+esc(a.CLIENTE||'Client')+'</div><div style="font-size:9.5px;color:var(--t3)">'+a.REFNUMPERS+'</div></div></div><span id="v360arr_'+a.IDOPOACT+'" style="font-size:10px;color:#8B5CF6;font-weight:600">▼ Desplegar</span>';

                var v360body = document.createElement('div');
                v360body.style.cssText = 'display:none;flex-direction:column';

                var v360tabBar = document.createElement('div');
                v360tabBar.style.cssText = 'display:flex;gap:20px;padding:10px 16px;background:var(--bg3);border-bottom:1px solid var(--br);justify-content:center;flex-wrap:wrap';

                // Contenidor clip (overflow:hidden, fons fosc per cross-origin)
                var v360clip = document.createElement('div');
                v360clip.style.cssText = 'position:relative;overflow:hidden;width:100%;height:370px;background:#16162a';

                // Placeholder (div separat, no srcdoc a l'iframe)
                var v360ph = document.createElement('div');
                v360ph.style.cssText = 'position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;color:#7070a0;font-size:12px;font-family:sans-serif;pointer-events:none';
                v360ph.innerHTML = '<span style="font-size:22px;opacity:.5">🔍</span><span>Fes clic a ▼ Desplegar per carregar</span>';

                // Iframe pur (sense srcdoc, evita conflicte)
                // clip=280: amaga nav (~50px) + header client (~80px) + logo/nom (~150px) → mostra NIF/adreça
                var v360clipPx = 280;
                var v360iframe = document.createElement('iframe');
                v360iframe.style.cssText = 'position:absolute;left:0;top:-'+v360clipPx+'px;width:100%;height:'+(370+v360clipPx)+'px;border:none;display:none';
                v360iframe.setAttribute('scrolling','yes');
                v360iframe.addEventListener('load', function(){
                    // Amagar placeholder quan l'iframe carregui (sigui X-Frame o no)
                    v360ph.style.opacity='0';
                    setTimeout(function(){v360ph.style.display='none';},300);
                });

                v360clip.appendChild(v360ph);
                v360clip.appendChild(v360iframe);

                var v360loaded = false;
                var activeV360Tab = null;

                function v360DoLoad(){
                    if(v360loaded) return;
                    v360loaded = true;
                    v360ph.innerHTML = '<span style="font-size:22px;opacity:.5">⏳</span><span>Carregant Visió 360...</span>';
                    v360iframe.style.display = 'block';
                    v360iframe.src = v360BaseUrl;
                }

                var tabDefs = [
                    {l:'👤 Fitxa',     visH:370, clip:280},  // salta nav+header+logo → mostra NIF/adreça/email
                    {l:'📄 Pòlisses',  visH:560, clip:50},   // salta nav → mostra tabs V360 + contingut
                    {l:'💳 Rebuts',    visH:560, clip:50},
                    {l:'⚠️ Sinistres', visH:560, clip:50}
                ];
                tabDefs.forEach(function(tab){
                    var tbtn = document.createElement('button');
                    tbtn.style.cssText = 'padding:6px 14px;border-radius:20px;border:1px solid var(--br);background:var(--bg2);color:var(--t3);font-size:12px;font-weight:600;cursor:pointer;font-family:inherit;transition:all .12s';
                    tbtn.textContent = tab.l;
                    tbtn.addEventListener('click', function(){
                        if(activeV360Tab){activeV360Tab.style.background='var(--bg2)';activeV360Tab.style.color='var(--t3)';activeV360Tab.style.borderColor='var(--br)';}
                        tbtn.style.background='rgba(139,92,246,.12)';tbtn.style.color='#8B5CF6';tbtn.style.borderColor='rgba(139,92,246,.3)';
                        activeV360Tab = tbtn;
                        v360clip.style.height = tab.visH+'px';
                        v360iframe.style.top = '-'+tab.clip+'px';
                        v360iframe.style.height = (tab.visH+tab.clip)+'px';
                        v360DoLoad();
                    });
                    v360tabBar.appendChild(tbtn);
                });

                v360body.appendChild(v360tabBar);
                v360body.appendChild(v360clip);

                var v360open = false;
                v360hdr.addEventListener('click', function(){
                    v360open = !v360open;
                    v360body.style.display = v360open ? 'flex' : 'none';
                    if(v360open) v360body.style.flexDirection = 'column';
                    v360hdr.style.borderBottomWidth = v360open ? '1px' : '0px';
                    var arr = SR.getElementById('v360arr_'+a.IDOPOACT);
                    if(arr) arr.textContent = v360open ? '▲ Col·lapsar' : '▼ Desplegar';
                    if(v360open) v360DoLoad();
                });

                cardV360.appendChild(v360hdr);
                cardV360.appendChild(v360body);
                cnt.appendChild(cardV360);
            }

            // ── CARD 2: Observacions (sempre visible si existeix) ────────────────
            if (obsTxt && obsTxt !== 'Sin detalles.') {
                var card2 = document.createElement('div');
                card2.style.cssText = 'background:var(--bg2);border:1px solid var(--br);border-radius:10px;padding:16px';
                card2.innerHTML = '<div style="font-size:11px;font-weight:700;color:var(--t3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px">💬 Observaciones</div>'+
                    '<div style="white-space:pre-wrap;line-height:1.6;font-size:12.5px;color:var(--t2)">'+esc(obsTxt)+'</div>';
                cnt.appendChild(card2);
            }

            // ── CARD 3: Notas ──────────────────────────────────────────────
            var card3 = document.createElement('div');
            card3.id = 'det-notes-'+a.IDOPOACT;
            card3.style.cssText = 'background:var(--bg2);border:1px solid var(--br);border-radius:10px;padding:16px';
            card3.innerHTML = '<div style="font-size:11px;font-weight:700;color:var(--t3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px">📎 Notas</div>'+
                '<div style="display:flex;align-items:center;gap:8px;color:var(--t3);font-size:11px"><div class="sp"></div>Carregant...</div>';
            cnt.appendChild(card3);

            // ── CARD 4: Activitats relacionades ───────────────────────────
            var relatedActs = S.all.filter(function(x){ return !x._isOpo && x.IDOPOACT === a.IDOPOACT; });
            var card4 = document.createElement('div');
            card4.style.cssText = 'background:var(--bg2);border:1px solid var(--br);border-radius:10px;padding:16px';
            var c4h = '<div style="font-size:11px;font-weight:700;color:var(--t3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px">🗓️ Activitats relacionades <span style="background:var(--bg3);padding:1px 7px;border-radius:10px;font-size:10px;font-weight:700;color:var(--t2)">'+relatedActs.length+'</span></div>';
            if (relatedActs.length === 0) {
                c4h += '<div style="font-size:11px;color:var(--t4);font-style:italic">Cap activitat vinculada. <a href="#" onclick="event.preventDefault();window.openCRM(\'ISC.Gaan.Oportunidad.FE.DetalleOportunidad\',\'IDOPOACT\',\''+a.IDOPOACT+'\')" style="color:#3B82F6;text-decoration:none">Crear a Gestiona Original ↗</a></div>';
            } else {
                c4h += '<div style="display:flex;flex-direction:column;gap:5px">';
                relatedActs.sort(function(a,b){return(b.TIM_INICI||'').localeCompare(a.TIM_INICI||'');}).forEach(function(ra){
                    var raT = gT(ra); var raS = gS(ra); var raTi = TYPES[raT]||TYPES['Tarea'];
                    var stCol = raS==='co'?'#10B981':raS==='ca'?'var(--t4)':'#FF7A95';
                    c4h += '<div style="padding:7px 10px;background:var(--bg3);border-radius:7px;border-left:3px solid '+raTi.c+';display:grid;grid-template-columns:auto 1fr auto auto;gap:8px;align-items:center">'+
                        '<span style="font-size:13px">'+raTi.i+'</span>'+
                        '<span style="font-size:12px;font-weight:600;color:var(--t1)">'+esc(ra.DESASUN||'—')+'</span>'+
                        '<span style="font-size:10px;color:var(--t3)">'+((ra.TIM_INICI||'').slice(0,10))+'</span>'+
                        '<span style="font-size:10px;font-weight:700;color:'+stCol+'">'+gSL(raS)+'</span>'+
                    '</div>';
                });
                c4h += '</div>';
                c4h += '<div style="margin-top:8px;font-size:10px;color:var(--t4)"><a href="#" onclick="event.preventDefault();window.openCRM(\'ISC.Gaan.Oportunidad.FE.DetalleOportunidad\',\'IDOPOACT\',\''+a.IDOPOACT+'\')" style="color:#3B82F6;text-decoration:none">Gestionar a Gestiona Original ↗</a></div>';
            }
            card4.innerHTML = c4h;
            cnt.appendChild(card4);

        } else {
            // ── ACTIVITAT: Card dades + observacions ──────────────────────
            var cardA = document.createElement('div');
            cardA.style.cssText = 'background:var(--bg2);border:1px solid var(--br);border-radius:10px;padding:16px';
            var dA = '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px">';
            dA += '<div><div style="font-size:9.5px;color:var(--t3);text-transform:uppercase;font-weight:700;margin-bottom:3px">👤 Client</div><div style="font-size:13px;font-weight:600;color:var(--t1)">'+(a.CLIENTE||'—')+'</div></div>';
            dA += '<div><div style="font-size:9.5px;color:var(--t3);text-transform:uppercase;font-weight:700;margin-bottom:3px">📊 Estat</div><span style="display:inline-block;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:700;background:rgba(220,0,40,.1);color:#FF7A95">'+gSL(st)+'</span></div>';
            dA += '<div><div style="font-size:9.5px;color:var(--t3);text-transform:uppercase;font-weight:700;margin-bottom:3px">🛡️ Mediador</div><div style="font-size:12px;color:var(--t1)">'+uname+'</div></div>';
            dA += '<div><div style="font-size:9.5px;color:var(--t3);text-transform:uppercase;font-weight:700;margin-bottom:3px">⚡ Prioritat</div><div class="apr '+(ti.c||'')+'" style="display:inline-block;margin:0">'+(ti.l||'')+'</div></div>';
            dA += '</div>';
            cardA.innerHTML = dA;
            cnt.appendChild(cardA);

            if (obsTxt && obsTxt !== 'Sin detalles.') {
                var cardObs = document.createElement('div');
                cardObs.style.cssText = 'background:var(--bg2);border:1px solid var(--br);border-radius:10px;padding:16px';
                cardObs.innerHTML = '<div style="font-size:11px;font-weight:700;color:var(--t3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px">💬 Observaciones</div>'+
                    '<div style="white-space:pre-wrap;line-height:1.6;font-size:12.5px;color:var(--t2)">'+esc(obsTxt)+'</div>';
                cnt.appendChild(cardObs);
            }
        }
        box.appendChild(cnt);
        
        var ftr = document.createElement('div');
        ftr.style.cssText = 'padding:16px 20px;border-top:1px solid var(--br);background:var(--bg2);display:flex;justify-content:space-between';
        ftr.innerHTML = '<div style="display:flex;gap:12px">' +
            (a.REFNUMPERS ? '<button id="det-v360-' + a.IDACTIV + '" style="padding:8px 16px;background:var(--bg3);border:1px solid var(--br);border-radius:6px;color:var(--t1);font-size:12px;cursor:pointer;font-weight:600">🔍 Ver Ficha 360</button>' : '') +
            '</div>' +
            '<div style="display:flex;gap:12px">' +
            '<button id="det-open-' + a.IDACTIV + '" style="padding:8px 16px;background:#3B82F6;border:none;border-radius:6px;color:#fff;font-size:12px;cursor:pointer;font-weight:600">↗️ Abrir en Gestiona Original</button>' +
            '</div>';
        box.appendChild(ftr);
        ov.appendChild(box);
        app.appendChild(ov);
        setTimeout(function() {
            var cls = ov.querySelector('#det-close-' + a.IDACTIV);
            if (cls) cls.addEventListener('click', function() { ov.remove(); });
            var op = ov.querySelector('#det-open-' + a.IDACTIV);
            if (op) op.addEventListener('click', function() {
                if (isOpo) {
                    window.openCRM('ISC.Gaan.Oportunidad.FE.DetalleOportunidad', 'IDOPOACT', a.IDOPOACT);
                } else {
                    window.openCRM('ISC.Gaan.Actividad.FE.DetalleActividad', 'IDACTIV', a.IDACTIV);
                }
            });
            // Kanban select listener
            var kbSel = ov.querySelector('#kb-sel-'+a.IDOPOACT);
            if(kbSel && isOpo) {
                kbSel.addEventListener('change', function(){
                    localStorage.setItem('kb_'+a.IDOPOACT, this.value);
                    if(S.view==='kanban') applyF();
                });
            }
            // Mediador select listener
            var medSel = ov.querySelector('#med-sel-'+a.IDOPOACT);
            if(medSel && isOpo) {
                medSel.addEventListener('change', function(){
                    var newUid = this.value;
                    if(USERS[newUid]){
                        a._uid = newUid;
                        var asg=loadAsg(); asg[String(a.IDOPOACT)]=newUid; saveAsg(asg);
                        updC(); render();
                    }
                });
            }
            // Cierre Estimado date input
            var cierreInp = ov.querySelector('#cierre-inp-'+a.IDOPOACT);
            if(cierreInp && isOpo) {
                cierreInp.addEventListener('change', function(){
                    var p=this.value.split('-');
                    if(p.length===3){ a.TIM_FI=p[2]+'/'+p[1]+'/'+p[0]; }
                });
            }
            var proxInp = ov.querySelector('#prox-inp-'+a.IDOPOACT);
            if(proxInp && isOpo) {
                proxInp.addEventListener('change', function(){
                    var p = this.value.split('-');
                    if(p.length===3){
                        var nd = p[2]+'/'+p[1]+'/'+p[0];
                        a.TIM_INICI = nd;
                        _saveOverride(String(a.IDOPOACT), nd);
                        applyF();
                    }
                });
            }
            var v360 = ov.querySelector('#det-v360-' + a.IDACTIV);
            if (v360) v360.addEventListener('click', function() {
                window.openCRM('PER.Clientes.V360.FE.V360', 'IdReferencia', a.REFNUMPERS.trim());
            });
            // Carregar notes OPO de forma async
            if (isOpo) {
                var notesEl = ov.querySelector('#det-notes-' + a.IDOPOACT);
                if (notesEl) {
                    // Carregar notes OPO de forma async via then-chaining
                var notesEl2 = notesEl;
                var idopo2 = a.IDOPOACT;
                fetch('/FE/ISC.Gaan.CRM.FE.Inicio/?_CIA=SCO')
                .then(function(rr){return rr.text();})
                .then(function(agH){
                    var mmc=agH.match(/name="_GAAN_ClientContext_"[^>]*value='([^']+)'/);
                    if(!mmc)mmc=agH.match(/name="_GAAN_ClientContext_"[^>]*value="([^"]+)"/);
                    if(!mmc){if(notesEl2)notesEl2.innerHTML='<div style="font-size:11px;font-weight:700;color:var(--t3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px">📎 Notas</div><div style="font-size:11px;color:var(--t4)">Context no disponible.</div>';return;}
                    var ctx2=JSON.parse(mmc[1].replace(/&quot;/g,'"').replace(/&amp;/g,'&'));
                    // Obtenir csid correcte via GET de DetOportunidad (no via Inicio que retorna HTML de l'Agenda)
                    return fetch('/FE/ISC.Gaan.Oportunidad.FE.DetOportunidad/?_CIA=SCO&IDOPOACT='+idopo2)
                    .then(function(dr){return dr.text();})
                    .then(function(dhtml){
                        var mm3=dhtml.match(/name="_GAAN_ClientContext_"[^>]*value='([^']+)'/);
                        if(!mm3)mm3=dhtml.match(/name="_GAAN_ClientContext_"[^>]*value="([^"]+)"/);
                        var ctx3=ctx2;
                        if(mm3){try{ctx3=JSON.parse(mm3[1].replace(/&quot;/g,'"').replace(/&amp;/g,'&'));console.log('Notes: nou csid',ctx3.csid);}catch(e){}}
                        else{console.log('Notes: csid no extret de DetOportunidad, usant ctx2');}
                        return fetch('/FE/ISC.Gaan.Anotacion.FE.GridAnotacion/',{
                            method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},
                            body:new URLSearchParams({'_GAAN_Mensaje_Pantalla_':JSON.stringify({GridResultados:{},Cod_Mediador_Usuario:'17167',IDOPOACT:idopo2,IdReferente:idopo2}),'_GAAN_Comando_':'buscar','_GAAN_ClientContext_':JSON.stringify(Object.assign({},ctx3,{init:'ISC.Gaan.Anotacion.FE.GridAnotacion'})),'_GAAN_SPA_':'true'}).toString()
                        });
                    }).then(function(r2){return r2.json();})
                    .then(function(j){
                        console.log('Notes raw:',JSON.stringify(j).substring(0,400));
                        var notes=(j.message&&j.message.ListaResultados&&j.message.ListaResultados.FuenteDatos)||(j.message&&j.message.FuenteDatos)||[];
                        if(!notesEl2)return;
                        if(notes.length===0){
                            notesEl2.innerHTML='<div style="font-size:11px;font-weight:700;color:var(--t3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px">📎 Notas</div><div style="font-size:11px;color:var(--t4);font-style:italic">Cap nota. Obre a Gestiona Original per gestionar notes.</div>';
                            return;
                        }
                        var nh='<div style="font-size:11px;font-weight:700;color:var(--t3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px">📎 Notas <span style="background:var(--bg3);padding:1px 7px;border-radius:10px;font-size:10px;color:var(--t2)">'+notes.length+'</span></div>';
                        nh+='<div style="display:flex;flex-direction:column;gap:6px">';
                        notes.forEach(function(n){
                            var titulo=n.TITULO||n.DS_TITULO||n.Titulo||'—';
                            var desc=n.DESCRIPCION||n.DS_DESCRIPCION||n.Descripcion||'';
                            var fecha=n.FECCRE||n.FECHA||n.FechaCreacion||'';
                            var filename=n.FILENAME||n.FileName||n.DS_FILENAME||'';
                            var fileurl=n.URLFILE||n.DS_URLFILE||'';
                            nh+='<div style="padding:10px 12px;background:var(--bg3);border-radius:7px;border:1px solid var(--br)">'+
                                '<div style="display:flex;justify-content:space-between;margin-bottom:4px">'+
                                '<div style="font-size:12px;font-weight:600;color:var(--t1)">'+esc(titulo)+'</div>'+
                                '<div style="font-size:10px;color:var(--t3)">'+esc(fecha)+'</div></div>'+
                                (desc?'<div style="font-size:11px;color:var(--t2);margin-bottom:6px">'+esc(desc)+'</div>':'')+
                                (filename?'<a href="'+fileurl+'" target="_blank" style="font-size:11px;color:#3B82F6;text-decoration:none">📄 '+esc(filename)+'</a>':'')+
                                '</div>';
                        });
                        nh+='</div>';
                        notesEl2.innerHTML=nh;
                    });
                }).catch(function(e){console.log('Notes err:',e);if(notesEl2)notesEl2.innerHTML='<div style="font-size:11px;font-weight:700;color:var(--t3);text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px">📎 Notas</div><div style="font-size:11px;color:var(--t4);font-style:italic">Notes no carregades. <a href="#" onclick="event.preventDefault();window.openCRM(\'ISC.Gaan.Oportunidad.FE.DetalleOportunidad\',\'IDOPOACT\',\''+a.IDOPOACT+'\')" style="color:#3B82F6;text-decoration:none">Veure a Gestiona Original ↗</a></div>';});
                }
            }
        }, 50);
    } catch(e) { console.error('showDetail error:', e); }
}
function rKanban() {
    var wrap=q('cnt');if(!wrap)return;wrap.innerHTML='';
    var topBar = document.createElement('div');
    topBar.className = 'foc-bar'; 
    topBar.style.cssText = 'padding:8px 12px;background:var(--bg2);border-bottom:1px solid var(--br);display:flex;align-items:center;gap:12px;flex-shrink:0';
    topBar.innerHTML = '<span style="font-size:12px;font-weight:600;color:var(--t1)">Cierre Estimado:</span>';
    var kFils = [
       {k:'day', l:'Hoy'},
       {k:'week', l:'Esta Semana'},
       {k:'month', l:'Este Mes'},
       {k:'year', l:'Este Año'},
       {k:'all', l:'Cualquiera'}
    ];
    var eGrp = document.createElement('div');
    eGrp.className = 'foc-type';
    kFils.forEach(function(f) {
        var p = document.createElement('button');
        p.className = 'foc-tpill' + (S.kFoc === f.k ? ' on' : '');
        p.textContent = f.l;
        p.onclick = function() { S.kFoc = f.k; rKanban(); };
        eGrp.appendChild(p);
    });
    topBar.appendChild(eGrp);
    wrap.appendChild(topBar);

    var kw=document.createElement('div');
    kw.className='kb-wrap';
    kw.style.cssText='display:flex;height:100%;overflow-x:auto;padding:12px;gap:12px;background:var(--bg1)';
    
    var K_COLS = [
        {id: 'espera', name: 'En Espera', color: '#6B7280'},
        {id: 'analisis', name: 'En Análisis', color: '#3B82F6'},
        {id: 'tarifa', name: 'En Tarificación', color: '#F59E0B'},
        {id: 'presentar', name: 'Para Presentar', color: '#8B5CF6'},
        {id: 'respuesta', name: 'Pdte. Respuesta', color: '#EC4899'},
        {id: 'emitir', name: 'Para Emitir', color: '#10B981'},
        {id: 'cierre', name: 'Cierre o P. Año', color: '#059669', dropClose: true}
    ];
    
    var KCOLS_MAP = {};
    K_COLS.forEach(function(c) { KCOLS_MAP[c.id] = c; });
    var kbData = {};
    K_COLS.forEach(function(c) { kbData[c.id] = []; });
    
    var pad = function(n){return String(n).padStart(2,'0');};
    var todayStr = String(new Date().getDate()).padStart(2,'0')+'/'+String(new Date().getMonth()+1).padStart(2,'0')+'/'+new Date().getFullYear();
    var tod = new Date();
    var curMoStr = pad(tod.getMonth()+1)+'/'+tod.getFullYear();
    var curYrStr = String(tod.getFullYear());

    var filteredOpos = S.fil.filter(function(a){return a._isOpo;}).filter(function(a) {
        if(S.kFoc === 'all') return true;
        var fi = (a.TIM_FI||'').slice(0,10); 
        if(S.kFoc === 'day') return fi === todayStr;
        if(S.kFoc === 'month') return fi.length === 10 && fi.slice(3) === curMoStr;
        if(S.kFoc === 'year') return fi.length === 10 && fi.slice(6) === curYrStr;
        return true;
    });

    filteredOpos.forEach(function(a) {
        var dsArray = (a.TIM_INICI||'').split('/');
        var isPast = false;
        if(dsArray.length===3) {
            var ts = new Date(dsArray[2], dsArray[1]-1, dsArray[0]).getTime();
            var tmp = new Date();
            tmp.setHours(0,0,0,0);
            isPast = ts < tmp.getTime();
        }
        var colId = localStorage.getItem('kb_' + a.IDOPOACT);
        if(!colId || !KCOLS_MAP[colId]) {
            colId = isPast ? 'espera' : 'analisis';
        }
        if(colId === 'espera' && !isPast) {
            colId = 'analisis'; 
        }
        kbData[colId].push(a);
    });
    
    K_COLS.forEach(function(col) {
        var colDiv = document.createElement('div');
        colDiv.className='kb-col';
        colDiv.style.cssText='display:flex;flex-direction:column;min-width:280px;width:280px;background:var(--bg2);border-radius:8px;border:1px solid var(--br);overflow:hidden';
        
        var hdr = document.createElement('div');
        hdr.style.cssText='padding:10px 12px;font-weight:700;font-size:12px;display:flex;justify-content:space-between;align-items:center;border-bottom:2px solid ' + col.color;
        hdr.innerHTML = '<span>'+col.name+'</span><span style="font-size:10px;background:var(--bg3);padding:2px 6px;border-radius:10px">'+kbData[col.id].length+'</span>';
        colDiv.appendChild(hdr);
        
        var body = document.createElement('div');
        body.className = 'kb-body';
        body.dataset.id = col.id;
        body.style.cssText='flex:1;overflow-y:auto;padding:8px;display:flex;flex-direction:column;gap:8px';
        
        mkDrop(body, function(drag){
            var a=S.all.find(function(x){return x.IDACTIV===drag.id;});
            if(!a || !a._isOpo) return;
            localStorage.setItem('kb_' + a.IDOPOACT, col.id);
            if(col.dropClose && !gS(a).includes('co') && !gS(a).includes('ca')) {
               setTimeout(function(){
                   var ov=document.createElement('div');
                   ov.style.cssText='position:absolute;inset:0;background:rgba(0,0,0,.7);display:flex;align-items:center;justify-content:center;z-index:9999';
                   var box=document.createElement('div');
                   box.style.cssText='background:var(--bg2);padding:24px;border-radius:12px;width:320px;display:flex;flex-direction:column;gap:12px;border:1px solid var(--br);box-shadow:0 10px 30px rgba(0,0,0,0.5)';
                   box.innerHTML='<h3 style="margin:0;color:var(--t1);font-family:system-ui">Formalizar</h3><p style="font-size:12px;color:var(--t2)">'+(a.DESASUN||'').substring(0,40)+'</p>';
                   
                   box.innerHTML += '<button onclick="localStorage.setItem(\'kb_\'+'+a.IDOPOACT+', \'cierre\');_OAv.ca(\''+a.IDACTIV+'\');this.parentNode.parentNode.remove();applyF();" style="padding:12px;background:#10B981;color:#fff;border:none;border-radius:6px;cursor:pointer">🏆 Completada Ganada</button>';
                   box.innerHTML += '<button onclick="localStorage.setItem(\'kb_\'+'+a.IDOPOACT+', \'cierre\');_OAv.da(\''+a.IDACTIV+'\');this.parentNode.parentNode.remove();applyF();" style="padding:12px;background:#DC0028;color:#fff;border:none;border-radius:6px;cursor:pointer">❌ Descartada / Perdida</button>';
                   box.innerHTML += '<button onclick="localStorage.setItem(\'kb_\'+'+a.IDOPOACT+', \'espera\');this.parentNode.parentNode.remove();applyF();" style="padding:12px;background:transparent;color:var(--t3);border:none;cursor:pointer">Cancelar</button>';
                   
                   ov.appendChild(box);
                   document.body.appendChild(ov);
               }, 50);
            }
            applyF(); 
        }, false);
        
        kbData[col.id].sort((a,b)=>((a.TIM_INICI||'').split('/').reverse().join('').localeCompare((b.TIM_INICI||'').split('/').reverse().join(''))));
        
        kbData[col.id].forEach(function(a) {
            var c = document.createElement('div');
            var u = USERS[a._uid||a.UID_RACF||'M441819E']||USERS['M441819E'];
            var pr = gP(a);
            c.className = 'kb-card';
            c.style.cssText = 'background:var(--bg1);border:1px solid var(--br);border-radius:6px;padding:10px;font-size:11px;cursor:pointer;position:relative;border-left:3px solid ' + col.color + ';';
            c.innerHTML = '<div style="position:absolute;top:-8px;right:-8px;width:18px;height:18px;background:'+u.bg+';color:'+u.color+';border-radius:50%;font-size:8px;font-weight:700;display:grid;place-items:center;border:2px solid var(--bg1)">'+u.ini+'</div>' +
            '<div style="font-weight:600;font-size:12px;color:var(--t1);margin-bottom:4px;width:90%">'+(a.DESASUN||'Sin asunto').substring(0,60)+'</div>'+
            '<div style="color:var(--t2);margin-bottom:2px">'+(a.CLIENTE||'Sin cliente')+'</div>'+
            '<div style="color:var(--t4);font-size:9.5px;margin-bottom:4px">🎯 Cierre: '+(a.TIM_FI||'--/--/----')+'</div>';
            mkDrag(c, a.IDACTIV, false);
            c.onclick = function(e){ e.stopPropagation(); showDetail(a.IDACTIV); };
            body.appendChild(c);
        });
        
        colDiv.appendChild(body);
        kw.appendChild(colDiv);
    });
    
    wrap.appendChild(kw);
}
function rFocus(){var wrap=q('cnt');if(!wrap)return;wrap.innerHTML='';var fw=document.createElement('div');fw.className='foc-wrap';var bar=document.createElement('div');bar.className='foc-bar';var ulabel=document.createElement('span');ulabel.className='foc-label';ulabel.textContent='Usuario';bar.appendChild(ulabel);var usr=document.createElement('div');usr.className='foc-usr';UIDS.forEach(function(uid){var u=USERS[uid];var pill=document.createElement('button');pill.className='foc-upill'+(S.foc.uid===uid?' on':'');pill.style.borderColor=S.foc.uid===uid?u.color:'transparent';if(S.foc.uid===uid)pill.style.background=u.bg;pill.style.color=S.foc.uid===uid?u.color:'';pill.textContent=u.ini;pill.title=u.name;pill.addEventListener('click',function(){S.foc.uid=uid;S.foc.types=new Set();rFocus();});usr.appendChild(pill);});bar.appendChild(usr);var sep=document.createElement('div');sep.className='foc-sep';bar.appendChild(sep);var eLabel=document.createElement('span');eLabel.className='foc-label';eLabel.textContent='Registro';bar.appendChild(eLabel);var eGrp=document.createElement('div');eGrp.className='foc-type';[{k:'all',l:'Todos',cls:''},{k:'opo',l:'💼 OPOs',cls:'on-opo'},{k:'act',l:'🗓 Actividades',cls:''}].forEach(function(e){var p=document.createElement('button');var isOn=S.foc.filter===e.k;p.className='foc-tpill'+(isOn?' on'+(e.cls?' '+e.cls:''):'');p.textContent=e.l;p.addEventListener('click',function(){S.foc.filter=e.k;S.foc.types=new Set();rFocus();});eGrp.appendChild(p);});bar.appendChild(eGrp);if(S.foc.filter!=='opo'){var sepD=document.createElement('div');sepD.className='foc-sep';bar.appendChild(sepD);var dToggle=document.createElement('button');dToggle.className='foc-tpill'+(S.foc.allDays?' on':'');dToggle.textContent=(S.foc.allDays?'📅 Todos los días':'📅 Solo '+focDate);dToggle.addEventListener('click',function(){S.foc.allDays=!S.foc.allDays;rFocus();});bar.appendChild(dToggle);var sep2=document.createElement('div');sep2.className='foc-sep';bar.appendChild(sep2);var tLabel=document.createElement('span');tLabel.className='foc-label';tLabel.textContent='Tipo';bar.appendChild(tLabel);var tGrp=document.createElement('div');tGrp.className='foc-type';var types=Object.keys(TYPES).filter(function(t){return t!=='Oportunidad';});types.forEach(function(t){var ti=TYPES[t];var p=document.createElement('button');var isOn=S.foc.types.has(t);p.className='foc-tpill'+(isOn?' on':'');if(isOn){p.style.background='rgba('+ti.c.slice(1).match(/../g).map(function(h){return parseInt(h,16);}).join(',')+',.12)';p.style.borderColor=ti.c;p.style.color=ti.c;}
p.innerHTML=ti.i+' '+t;p.addEventListener('click',function(){if(S.foc.types.has(t))S.foc.types.delete(t);else S.foc.types.add(t);rFocus();});tGrp.appendChild(p);});bar.appendChild(tGrp);}
fw.appendChild(bar);var list=document.createElement('div');list.className='foc-list';var focDate=String(S.date.getDate()).padStart(2,'0')+'/'+String(S.date.getMonth()+1).padStart(2,'0')+'/'+S.date.getFullYear();var items=S.fil.filter(function(a){var uid=a._uid||a.UID_RACF||'M441819E';if(uid!==S.foc.uid)return false;var tim=(a.TIM_INICI||'').slice(0,10);if(!S.foc.allDays&&tim&&tim!==focDate)return false;var isOpo=a._isOpo||gT(a)==='Oportunidad';if(S.foc.filter==='opo'&&!isOpo)return false;if(S.foc.filter==='act'&&isOpo)return false;if(S.foc.types.size>0&&!S.foc.types.has(gT(a)))return false;return true;});if(items.length===0){var empty=document.createElement('div');empty.className='foc-empty';empty.innerHTML='<div class="foc-empty-ico">📋</div><div class="foc-empty-txt">Sin registros para '+focDate+'</div>';list.appendChild(empty);}else{items.sort(function(a,b){var ta=a.TIM_INICI_HORA||'';var tb=b.TIM_INICI_HORA||'';return ta<tb?-1:ta>tb?1:0;});items.forEach(function(a){var fc=document.createElement('div');var isOpo=a._isOpo||gT(a)==='Oportunidad';var st=gS(a);var ti=gP(a);fc.className='foc-card'+(isOpo?' foc-opo':'')+(st!=='ab'?' sdone':'');fc.style.borderLeftColor=isOpo?'#8B5CF6':ti.c;var t1=a.TIM_INICI_HORA||'';var t2=a.TIM_FI_HORA||'';var timeStr=t1?(t1+(t2&&t2!==t1?'–'+t2:'')):'Sin hora';fc.innerHTML='<div class="foc-top">'+
'<span class="foc-time">'+timeStr+'</span>'+
'<span class="foc-subj">'+esc(a.DESASUN||'Sin asunto')+'</span>'+
'<span class="foc-st '+st+'">'+gSL(st)+'</span>'+
'</div>'+
(a.CLIENTE?'<div class="foc-cli">👤 '+esc(a.CLIENTE)+'</div>':'')+
'<div class="foc-meta">'+
'<span class="foc-badge">'+ti.i+' '+gT(a)+'</span>'+
(a.RAMEMIS_DESC?'<span class="foc-badge">'+esc(a.RAMEMIS_DESC)+'</span>':'')+
'</div>';fc.addEventListener('click',function(){_OAv.showDetail(a,fc);});list.appendChild(fc);});}
fw.appendChild(list);wrap.appendChild(fw);}
function rTimeline(){var wrap=q('cnt');if(!wrap)return;wrap.innerHTML='';var tl=document.createElement('div');tl.className='tl-wrap';var today=new Date();var days=[];var start=new Date(today);start.setDate(start.getDate()-7);for(var i=0;i<60;i++){var d=new Date(start);d.setDate(start.getDate()+i);days.push(d);}
UIDS.forEach(function(uid){var u=USERS[uid];if(!S.aU.has(uid))return;var userItems=S.fil.filter(function(a){return(a._uid||a.UID_RACF||'M441819E')===uid;});var section=document.createElement('div');section.className='tl-user';var uhdr=document.createElement('div');uhdr.className='tl-uhdr';uhdr.innerHTML='<div class="av" style="width:22px;height:22px;background:'+u.bg+';color:'+u.color+';font-size:8px;font-weight:700;border-radius:50%;display:grid;place-items:center;flex-shrink:0">'+u.ini+'</div>'+
'<span class="tl-uname">'+u.name+'</span><span class="tl-cnt">'+userItems.length+' registros</span>';section.appendChild(uhdr);var scroll=document.createElement('div');scroll.className='tl-scroll';var track=document.createElement('div');track.className='tl-track';days.forEach(function(day){var iso=fmtISO(day);var dd=String(day.getDate()).padStart(2,'0');var mm=['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'][day.getMonth()];var isToday=iso===fmtISO(today);var dayISO=String(day.getDate()).padStart(2,'0')+'/'+String(day.getMonth()+1).padStart(2,'0')+'/'+day.getFullYear();var dayItems=userItems.filter(function(a){return(a.TIM_INICI||'').slice(0,10)===dayISO;});var col=document.createElement('div');col.className='tl-day'+(isToday?' tl-today':'');var dhd=document.createElement('div');dhd.className='tl-dhd'+(isToday?' td':'');dhd.textContent=dd+'/'+mm;col.appendChild(dhd);if(dayItems.length===0){}else{dayItems.forEach(function(a){var item=document.createElement('div');var isOpo=a._isOpo;item.className='tl-item'+(isOpo?' tl-opo':'');if(!isOpo){var ti=gP(a);item.style.borderLeftColor=ti.c;}
item.textContent=esc(a.DESASUN||a.CLIENTE||'Sin asunto').slice(0,20);item.title=(a.DESASUN||'')+(a.CLIENTE?' — '+a.CLIENTE:'');item.addEventListener('click',function(){_OAv.showDetail(a,item);});col.appendChild(item);});}
track.appendChild(col);});scroll.appendChild(track);section.appendChild(scroll);tl.appendChild(section);setTimeout(function(){var todayCol=track.querySelectorAll('.tl-today')[0];if(todayCol)todayCol.scrollIntoView({inline:'center',block:'nearest'});},50);});wrap.appendChild(tl);}
function render(){if(S.view==='day')rDay();else if(S.view==='week')rWeek();else if(S.view==='month')rMonth();else if(S.view==='timeline')rTimeline();else if(S.view==='focus')rFocus();else if(S.view==='kanban')rKanban();else rList();}
function rDay(){try{var ds=fmtD(S.date);var sv=S.vtos.show;var activeUIDS=UIDS.filter(function(uid){return S.aU.has(uid);});var nc=sv?activeUIDS.length+1:activeUIDS.length;var wrap=document.createElement('div');wrap.style.cssText='flex:1;min-height:0;overflow:hidden;display:flex;flex-direction:column';var cols=document.createElement('div');cols.className='dcols';cols.style.cssText='grid-template-columns:'+(sv?'248px repeat('+activeUIDS.length+',1fr)':'repeat('+activeUIDS.length+',1fr)')+';flex:1';if(sv){var vc=document.createElement('div');vc.className='ucol';vc.style.borderRight='2px solid rgba(220,0,40,.3)';var vh=document.createElement('div');vh.className='uch';vh.innerHTML='<div class="av" style="width:30px;height:30px;border-radius:7px;background:rgba(86,86,106,.2);color:#DC0028;font-size:16px">📋</div>'+
'<div><div style="font-size:11px;font-weight:700;color:var(--t2)">OPO / VTOS</div><div style="font-size:9px;color:var(--t3)">'+S.vtos.items.length+' entradas</div></div>'+
'<button onclick="_OAv.vtosNew()" style="margin-left:auto;width:22px;height:22px;border-radius:5px;border:1px solid var(--br);background:transparent;color:var(--t3);font-size:14px;display:grid;place-items:center;cursor:pointer">+</button>';var vb=document.createElement('div');vb.className='ucb';mkDrop(vb,function(drag){if(drag.vtos)return;var a=S.all.find(function(x){return x.IDACTIV===drag.id;});if(!a)return;var it={id:Date.now()+'',asunto:(a.DESASUN||'Sin asunto'),cliente:(a.CLIENTE||''),fecha:(a.TIM_INICI||fmtD(S.date)),tipo:gT(a),notas:(a.OBSACT||''),creado:Date.now()};S.vtos.items.unshift(it);saveVtos(S.vtos.items);updVtos();render();},true);if(S.vtos.items.length===0){vb.innerHTML='<div class="empty"><div class="empi">📋</div><div class="empt">Sin entradas<br>Pulsa + para añadir</div></div>';}
else{S.vtos.items.forEach(function(it,idx){var va=document.createElement('div');va.className='vact';va.style.animationDelay=(idx*.03)+'s';va.innerHTML='<div class="vtit">'+esc(it.asunto||'—')+'</div>'+
'<div class="vcli2">'+esc(it.cliente||'—')+'</div>'+
'<div class="vfoot"><span class="vdate">'+(it.fecha||'')+'</span>'+
'<button class="vdel" onclick="event.stopPropagation();_OAv.vtosDelete(\''+it.id+'\')" title="Eliminar">×</button></div>';mkDrag(va,it.id,true);va.addEventListener('click',function(){_OAv.vtosEdit(it.id);});vb.appendChild(va);});}
vc.appendChild(vh);vc.appendChild(vb);cols.appendChild(vc);}
activeUIDS.forEach(function(uid){var u=USERS[uid];var acts=S.fil.filter(function(a){return(a._uid||a.UID_RACF||'M441819E')===uid&&(a.TIM_INICI||'').slice(0,10)===ds;}).sort(function(a,b){return(a.TIM_INICI_HORA||'').localeCompare(b.TIM_INICI_HORA||'');});var col=document.createElement('div');col.className='ucol';var hdr2=document.createElement('div');hdr2.className='uch';hdr2.innerHTML='<div class="av" style="width:30px;height:30px;border-radius:7px;background:'+u.bg+';color:'+u.color+';font-size:9.5px">'+u.ini+'</div><div><div style="font-size:11px;font-weight:600;color:var(--t1)">'+u.name+'</div><div style="font-size:9px;color:var(--t3)">'+acts.length+' act.</div></div>';var body=document.createElement('div');body.className='ucb';mkDrop(body,function(drag){if(drag.vtos){_OAv.vtosAssign(drag.id,uid);return;}
var a=S.all.find(function(x){return x.IDACTIV===drag.id;});if(!a)return;var pD=a.TIM_INICI;a._uid=uid;applyF();if(a.TIM_INICI&&a.TIM_INICI!==pD)reschedCRM(a.IDACTIV,a.TIM_INICI);},false);if(acts.length===0){var dsKey=ds.slice(6,10)+ds.slice(3,5)+ds.slice(0,2);var userActs=S.fil.filter(function(a){return(a._uid||a.UID_RACF||'M441819E')===uid&&(a.TIM_INICI||'').length>=10;});userActs.sort(function(a,b){var pa=a.TIM_INICI.slice(6,10)+a.TIM_INICI.slice(3,5)+a.TIM_INICI.slice(0,2),pb=b.TIM_INICI.slice(6,10)+b.TIM_INICI.slice(3,5)+b.TIM_INICI.slice(0,2);return pa>pb?1:-1;});var nxt=userActs.find(function(a){return(a.TIM_INICI.slice(6,10)+a.TIM_INICI.slice(3,5)+a.TIM_INICI.slice(0,2))>dsKey;});var hintStr=nxt?'<div style="font-size:9px;color:var(--t4);margin-top:4px">Pròxima: '+nxt.TIM_INICI.slice(0,10)+'</div>':'';body.innerHTML='<div class="empty"><div class="empi">📋</div><div class="empt">Sense activitats<br>'+(isToday(S.date)?'per avui':'per aquesta data')+'</div>'+hintStr+'</div>';}
else{acts.forEach(function(a,i){body.appendChild(card(a,i));});}
col.appendChild(hdr2);col.appendChild(body);cols.appendChild(col);});wrap.appendChild(cols);CNT.innerHTML='';CNT.appendChild(wrap);}catch(e){console.error('OA7 rDay error:',e);CNT.innerHTML='<div class="empty"><div class="empi">⚠️</div><div class="empt">Error vista Día: '+e.message+'</div></div>';}}
function rWeek(){var days=gWk(S.date),tod=fmtISO(new Date()),sel=fmtISO(S.date);var wrap=document.createElement('div');wrap.className='wkwrap';var hdr2=document.createElement('div');hdr2.className='wkhdr';var grid=document.createElement('div');grid.className='wkgrid';days.forEach(function(day,i){var ds=fmtD(day),iso=fmtISO(day),isT=iso===tod,isSel=iso===sel;var acts=S.fil.filter(function(a){return(a.TIM_INICI||'').slice(0,10)===ds;}).sort(function(a,b){return(a.TIM_INICI_HORA||'').localeCompare(b.TIM_INICI_HORA||'');});var dh=document.createElement('div');dh.className='wkdh'+(isT?' wkt':'')+(isSel?' wks':'');dh.innerHTML='<div class="wkdn">'+DNAMES[i]+'</div><div class="wkdd">'+day.getDate()+'</div><div class="wkdc">'+acts.length+' act.</div>';dh.addEventListener('click',function(){_OAv.jd(day.getFullYear(),day.getMonth(),day.getDate());});hdr2.appendChild(dh);var col=document.createElement('div');col.className='wkcol';mkDrop(col,function(drag){var a=S.all.find(function(x){return x.IDACTIV===drag.id;});if(!a)return;var pD=a.TIM_INICI;a.TIM_INICI=ds;applyF();if(a.TIM_INICI!==pD)reschedCRM(a.IDACTIV,a.TIM_INICI);},false);if(acts.length===0){var e=document.createElement('div');e.style.cssText='font-size:9px;color:var(--t4);padding:8px 4px;text-align:center';e.textContent='—';col.appendChild(e);}
else{acts.forEach(function(a,idx){var u=USERS[a._uid||a.UID_RACF||'M441819E']||USERS['M441819E'];var ti=TYPES[gT(a)]||TYPES['Tarea'];var wa=document.createElement('div');wa.className='wact'+(gS(a)==='co'?' dn':'');wa.style.cssText='animation-delay:'+(idx*.02)+'s;border-left-color:'+ti.c;var wusel=document.createElement('select');wusel.className='wusel';UIDS.forEach(function(uid2){var op=document.createElement('option');op.value=uid2;op.textContent=USERS[uid2].ini;if(uid2===(a._uid||a.UID_RACF||'M441819E'))op.selected=true;wusel.appendChild(op);});wusel.style.cssText='background:'+u.bg+';color:'+u.color+';border-color:'+u.color;wusel.addEventListener('mousedown',function(e){e.stopPropagation();});wusel.addEventListener('click',function(e){e.stopPropagation();});wusel.addEventListener('change',function(e){e.stopPropagation();_OAv.rua(a.IDACTIV,this.value);});wa.innerHTML='<div class="warow"><span class="wat">'+(a.TIM_INICI_HORA||'--:--')+'</span><span class="was">'+esc(a.DESASUN||'Sin asunto')+'</span></div>';wa.querySelector('.warow').appendChild(wusel);var wDel=document.createElement('button');wDel.className='wdel';wDel.innerHTML='🗑';wDel.title='Cancelar';wDel.addEventListener('mousedown',function(e){e.stopPropagation();});wDel.addEventListener('click',function(e){e.stopPropagation();_OAv.da(a.IDACTIV);});wa.querySelector('.warow').appendChild(wDel);wa.addEventListener('click',function(e){if(e.target.tagName==='SELECT'||e.target.classList.contains('wdel'))return;showDetail(a.IDACTIV);});mkDrag(wa,a.IDACTIV,false);col.appendChild(wa);});}
grid.appendChild(col);});wrap.appendChild(hdr2);wrap.appendChild(grid);CNT.innerHTML='';CNT.appendChild(wrap);}
function rMonth(){var days=gMo(S.date),tod=fmtISO(new Date()),sel=fmtISO(S.date);var wrap=document.createElement('div');wrap.className='mowrap';var mhdr=document.createElement('div');mhdr.className='mohdr';DNAMES.forEach(function(d){var dh=document.createElement('div');dh.className='modh';dh.textContent=d;mhdr.appendChild(dh);});var mgrid=document.createElement('div');mgrid.className='mogrid';days.forEach(function(day){var ds=fmtD(day),iso=fmtISO(day);var acts=S.fil.filter(function(a){return(a.TIM_INICI||'').slice(0,10)===ds;});var cell=document.createElement('div');cell.className='moday'+(iso===tod?' mt':'')+(iso===sel?' ms':'')+(day._o?' mo':'');cell.innerHTML='<div class="modn">'+day.getDate()+'</div>';if(acts.length>0){var dots=document.createElement('div');dots.className='modots';acts.slice(0,6).forEach(function(a){var dot=document.createElement('div');dot.className='modot';dot.style.background=(USERS[a._uid||a.UID_RACF||'M441819E']||USERS['M441819E']).color;dots.appendChild(dot);});if(acts.length>6){var mm=document.createElement('span');mm.style.cssText='font-size:9px;color:var(--t3)';mm.textContent='+'+(acts.length-6);dots.appendChild(mm);}
cell.appendChild(dots);acts.slice(0,2).forEach(function(a){var ti=TYPES[gT(a)]||TYPES['Tarea'];var lb=document.createElement('div');lb.className='molbl';lb.textContent=ti.i+' '+(a.DESASUN||'').substring(0,16);lb.style.cssText='border-left:2px solid '+ti.c+';padding-left:4px';cell.appendChild(lb);});if(acts.length>2){var mm=document.createElement('div');mm.className='molbl';mm.style.color='var(--t3)';mm.textContent='+'+(acts.length-2)+' más';cell.appendChild(mm);}}
cell.onclick=function(){_OAv.jd(day.getFullYear(),day.getMonth(),day.getDate());};mkDrop(cell,function(drag){var a=S.all.find(function(x){return x.IDACTIV===drag.id;});if(!a)return;var pD=a.TIM_INICI;a.TIM_INICI=ds;applyF();if(a.TIM_INICI!==pD)reschedCRM(a.IDACTIV,a.TIM_INICI);},false);mgrid.appendChild(cell);});wrap.appendChild(mhdr);wrap.appendChild(mgrid);CNT.innerHTML='';CNT.appendChild(wrap);}
function rList(){var avuiStr=fmtD(new Date());var setmana=gWk(new Date()).map(function(d){return fmtD(d);});var nowM=new Date();var sorted=[].concat(S.fil).filter(function(a){var dt=(a.TIM_INICI||'').slice(0,10);var okDate=true;if(S.listFil==='hoy')okDate=dt===avuiStr;else if(S.listFil==='semana')okDate=setmana.indexOf(dt)>=0;else if(S.listFil==='mes'){var p=dt.split('/');okDate=p[1]===pad(nowM.getMonth()+1)&&p[2]===String(nowM.getFullYear());}
var isOpo=a._isOpo||gT(a)==='Oportunidad';var okRt=!S.listRt||(S.listRt==='opo'&&isOpo)||(S.listRt==='act'&&!isOpo);var okTp=!S.listTp||gT(a)===S.listTp;return okDate&&okRt&&okTp;}).sort(function(a,b){var da=(a.TIM_INICI||'').split('/').reverse().join('-'),db=(b.TIM_INICI||'').split('/').reverse().join('-');return da.localeCompare(db)||((a.TIM_INICI_HORA||'').localeCompare(b.TIM_INICI_HORA||''));});var groups={};sorted.forEach(function(a){var k=a.TIM_INICI||'Sin fecha';(groups[k]||(groups[k]=[])).push(a);});var lv=document.createElement('div');lv.className='lview';
// Barra de filtres — tot en una sola fila
var fbar=document.createElement('div');fbar.style.cssText='display:flex;gap:4px;padding:4px 0 10px;flex-shrink:0;flex-wrap:wrap;align-items:center;';
// Separador visual
var sep=function(txt){var s=document.createElement('span');s.style.cssText='font-size:9px;color:var(--t4);margin:0 2px;';s.textContent=txt||'·';return s;};
// Filtres de data
[['','Tots'],['hoy','Avui'],['semana','Setmana'],['mes','Mes']].forEach(function(f){var btn=document.createElement('button');var isOn=S.listFil===f[0]||(f[0]===''&&!S.listFil);btn.className='rfb';if(isOn)btn.style.cssText='background:rgba(220,0,40,.12);border-color:rgba(220,0,40,.35);color:#DC0028;font-weight:600';btn.textContent=f[1];btn.addEventListener('click',function(){S.listFil=f[0]||null;render();});fbar.appendChild(btn);});
fbar.appendChild(sep());
// Filtres registre: OPO / ACT
[['','Tot'],['opo','💼 OPO'],['act','📋 ACT']].forEach(function(f){var btn=document.createElement('button');var isOn=S.listRt===f[0]||(f[0]===''&&!S.listRt);btn.className='rfb';if(isOn)btn.style.cssText='background:rgba(139,92,246,.12);border-color:rgba(139,92,246,.35);color:#8B5CF6;font-weight:600';btn.textContent=f[1];btn.addEventListener('click',function(){S.listRt=f[0]||null;render();});fbar.appendChild(btn);});
fbar.appendChild(sep());
// Filtres tipus activitat
Object.keys(TYPES).filter(function(t){return t!=='Oportunidad';}).forEach(function(t){var ti=TYPES[t];var btn=document.createElement('button');var isOn=S.listTp===t;btn.className='rfb';btn.title=t;btn.style.cssText=isOn?'background:rgba(220,0,40,.1);border-color:rgba(220,0,40,.3);color:var(--t1);font-weight:600':'';btn.textContent=ti.i;btn.addEventListener('click',function(){S.listTp=S.listTp===t?null:t;render();});fbar.appendChild(btn);});
lv.appendChild(fbar);
Object.entries(groups).forEach(function(e){var dg=document.createElement('div');dg.className='ldg';dg.textContent=e[0];lv.appendChild(dg);e[1].forEach(function(a){var ti=TYPES[gT(a)]||TYPES['Tarea'],u=USERS[a._uid||a.UID_RACF||'M441819E']||USERS['M441819E'],st=gS(a);var la=document.createElement('div');la.className='la';la.innerHTML='<span style="font-size:14px">'+ti.i+'</span><span style="font-size:10px;font-weight:600;color:var(--t2)">'+(a.TIM_INICI_HORA||'--:--')+'</span><span class="las">'+esc(a.DESASUN||'Sin asunto')+'</span><span class="lcli">'+esc(a.CLIENTE||'')+'</span><div id="lus_'+a.IDACTIV+'"></div><span class="ab '+st+'">'+gSL(st)+'</span>';var lusel=document.createElement('select');lusel.className='lusel';lusel.style.cssText='background:'+u.bg+';color:'+u.color+';border-color:'+u.color;UIDS.forEach(function(uid2){var op=document.createElement('option');op.value=uid2;op.textContent=USERS[uid2].ini;if(uid2===(a._uid||a.UID_RACF||'M441819E'))op.selected=true;lusel.appendChild(op);});lusel.addEventListener('mousedown',function(e){e.stopPropagation();});lusel.addEventListener('click',function(e){e.stopPropagation();});lusel.addEventListener('change',function(e){e.stopPropagation();_OAv.rua(a.IDACTIV,this.value);});var lph=la.querySelector('#lus_'+a.IDACTIV);if(lph)lph.replaceWith(lusel);var lDel=document.createElement('button');lDel.className='ldel';lDel.innerHTML='🗑';lDel.title='Cancelar actividad';lDel.addEventListener('mousedown',function(e){e.stopPropagation();});lDel.addEventListener('click',function(e){e.stopPropagation();_OAv.da(a.IDACTIV);});la.appendChild(lDel);la.addEventListener('click',function(e){if(e.target.tagName==='SELECT'||e.target.classList.contains('ldel'))return;showDetail(a.IDACTIV);});lv.appendChild(la);});});if(Object.keys(groups).length===0){var emDiv=document.createElement('div');emDiv.innerHTML='<div class="empty" style="margin-top:40px"><div class="empi">🔍</div><div class="empt">Sense resultats</div></div>';lv.appendChild(emDiv);}
CNT.innerHTML='';CNT.appendChild(lv);}
function vtosModal(it){q('vmodal')?.remove();var isE=!!(it&&it.id);var d=it||{id:'',asunto:'',cliente:'',fecha:fmtD(S.date),tipo:'Tarea',notas:''};var ov=document.createElement('div');ov.className='vmodal';ov.id='vmodal';var box=document.createElement('div');box.className='vmbox';box.innerHTML='<h3>'+(isE?'\u270f\ufe0f Editar entrada':'OPO/VTOS — Nueva entrada')+'</h3>'+
'<p>Items locales. Al asignar a un colaborador se abrir\u00e1 el CRM para crear la actividad.</p>'+
'<div class="vf"><label>Asunto *</label><input id="va" type="text" value="'+esc(d.asunto)+'"></div>'+
'<div class="vf"><label>Cliente</label><input id="vc" type="text" value="'+esc(d.cliente)+'"></div>'+
'<div class="vf"><label>Fecha (DD/MM/AAAA)</label><input id="vfe" type="text" value="'+esc(d.fecha)+'"></div>'+
'<div class="vf"><label>Tipo</label><select id="vti">'+Object.keys(TYPES).map(function(t){return'<option value="'+t+'"'+(d.tipo===t?' selected':'')+'>'+TYPES[t].i+' '+t+'</option>';}).join('')+'</select></div>'+
'<div class="vf"><label>Notas</label><textarea id="vn">'+esc(d.notas||'')+'</textarea></div>'+
'<div class="vmbtns" id="vmbtns"></div>';var btns=box.querySelector('#vmbtns');var bCancel=document.createElement('button');bCancel.className='vmcancel';bCancel.textContent='Cancelar';bCancel.addEventListener('click',function(){q('vmodal')?.remove();});var bSave=document.createElement('button');bSave.className='vmsave';bSave.textContent=isE?'Guardar':'A\xf1adir';bSave.addEventListener('click',function(){_OAv.vtosSave(d.id||'');});btns.appendChild(bCancel);btns.appendChild(bSave);ov.appendChild(box);app.appendChild(ov);setTimeout(function(){var el=q('va');if(el)el.focus();},50);}
window.openCRM = async function(mod, k, v) {
  try{
    var agPage = await fetch('/FE/ISC.Gaan.CRM.FE.Inicio/?_CIA=SCO');
    var agHtml = await agPage.text();
    var mm = agHtml.match(/name="_GAAN_ClientContext_"[^>]*value='([^']+)'/);
    if(!mm) mm = agHtml.match(/name="_GAAN_ClientContext_"[^>]*value="([^"]+)"/);
    if(!mm){alert('No s\'ha pogut obtenir el context del CRM');return;}
    var ctx = JSON.parse(mm[1].replace(/&quot;/g,'"').replace(/&amp;/g,'&'));

    var html;

    if(mod.indexOf('V360')>=0){
      var f = document.createElement('form');
      f.method='POST'; f.action='/FE/ISC.Gaan.CRM.FE.Inicio/'; f.target='_v360win'; f.style.display='none';
      [{n:'_GAAN_Mensaje_Pantalla_',v:JSON.stringify({IdCliente:v,Url:'vercliente'})},
       {n:'_GAAN_Comando_',v:'vercliente'},
       {n:'_GAAN_ClientContext_',v:JSON.stringify(ctx)}
      ].forEach(function(fi){var i=document.createElement('input');i.type='hidden';i.name=fi.n;i.value=fi.v;f.appendChild(i);});
      document.body.appendChild(f);
      // Obrir en segon pla (finestra petita fora de pantalla)
      var v360win = window.open('','_v360win','width=1,height=1,left=-9999,top=-9999');
      f.submit();
      setTimeout(function(){try{document.body.removeChild(f);}catch(e){}},300);
      // Tancar als 2.5 segons
      setTimeout(function(){try{if(v360win&&!v360win.closed)v360win.close();}catch(e){}}, 2500);
      return;
    }
      var cmd = mod.indexOf('Oportunidad')>=0 ? 'veropor' : 'veract';
      var msg = mod.indexOf('Oportunidad')>=0
        ? {IdOportunidad:v, Url:'veropor'}
        : {IdActividad:v, Url:'veract'};
      var r = await fetch('/FE/ISC.Gaan.CRM.FE.Inicio/', {
        method:'POST',
        headers:{'Content-Type':'application/x-www-form-urlencoded'},
        body: new URLSearchParams({
          '_GAAN_Mensaje_Pantalla_': JSON.stringify(msg),
          '_GAAN_Comando_': cmd,
          '_GAAN_ClientContext_': JSON.stringify(ctx)
        }).toString()
      });
      if(!r.ok){alert('Error '+r.status);return;}
      html = await r.text();

    var base = '<base href="https://grupoaplicaciones.gco.global/">';
    html = html.replace('<head>', '<head>'+base);
    var win = window.open('','_blank');
    win.document.open(); win.document.write(html); win.document.close();
  }catch(e){
    console.error('openCRM:',e); alert('Error: '+e.message);
  }
};
function showImportRpt(){var log=S.importLog||[];var acts=log.filter(function(x){return x.t==='act';});var opos=log.filter(function(x){return x.t==='opo';});var byUser={};UIDS.forEach(function(uid){byUser[uid]={acts:0,opos:0,types:{}};});log.forEach(function(x){var b=byUser[x.uid]||(byUser[x.uid]={acts:0,opos:0,types:{}});if(x.t==='act')b.acts++;else b.opos++;b.types[x.ta]=(b.types[x.ta]||0)+1;});var ov=document.createElement('div');ov.style.cssText='position:absolute;inset:0;background:rgba(0,0,0,.85);display:flex;align-items:flex-start;justify-content:center;z-index:10;overflow-y:auto;padding:24px 16px;';var box=document.createElement('div');box.style.cssText='background:var(--bg2);border:1px solid var(--brl);border-radius:14px;padding:24px;max-width:700px;width:100%;font-size:12px;line-height:1.5;';var h='<div style="font-weight:700;font-size:15px;margin-bottom:4px;color:var(--t1);font-family:inherit">📊 Informe d\'importació — Agenda v8</div>';h+='<div style="color:var(--t3);font-size:11px;margin-bottom:16px;font-family:inherit">'+new Date().toLocaleString('ca-ES')+'</div>';h+='<div style="display:flex;gap:12px;margin-bottom:20px;">';[['Activitats',acts.length,'#3B82F6'],['Oportunitats',opos.length,'#8B5CF6'],['Total',log.length,'#10B981']].forEach(function(c){h+='<div style="flex:1;background:var(--bg3);border-radius:8px;padding:12px;text-align:center;border:1px solid var(--br)"><div style="font-size:24px;font-weight:700;color:'+c[2]+'">'+c[1]+'</div><div style="color:var(--t3);font-size:11px;font-family:inherit">'+c[0]+'</div></div>';});h+='</div>';h+='<div style="font-weight:700;color:var(--t1);margin-bottom:8px;font-family:inherit;font-size:12px">Per usuari</div>';h+='<table style="width:100%;border-collapse:collapse;margin-bottom:20px;font-family:inherit">';h+='<tr style="color:var(--t3);border-bottom:1px solid var(--br);font-size:11px"><td style="padding:5px 8px">Usuari</td><td style="padding:5px 8px;text-align:center">Act.</td><td style="padding:5px 8px;text-align:center">Opos.</td><td style="padding:5px 8px;text-align:center">Total</td><td style="padding:5px 8px">Tipus (top 4)</td></tr>';UIDS.forEach(function(uid){var u=USERS[uid];var d=byUser[uid]||{acts:0,opos:0,types:{}};var top=Object.entries(d.types).sort(function(a,b){return b[1]-a[1];}).slice(0,4).map(function(e){return'<span style="background:var(--bg4);padding:1px 5px;border-radius:4px;margin-right:3px">'+e[0]+' <b>'+e[1]+'</b></span>';}).join('');h+='<tr style="border-bottom:1px solid var(--br)"><td style="padding:6px 8px"><span style="color:'+u.color+';font-weight:700;font-family:inherit">'+u.ini+'</span> <span style="color:var(--t1)">'+u.name+'</span></td><td style="padding:6px 8px;text-align:center;color:var(--t1);font-weight:700">'+d.acts+'</td><td style="padding:6px 8px;text-align:center;color:var(--t1);font-weight:700">'+d.opos+'</td><td style="padding:6px 8px;text-align:center;color:'+(d.acts+d.opos>0?u.color:'var(--t4)')+';font-weight:700">'+(d.acts+d.opos)+'</td><td style="padding:6px 8px;font-size:11px">'+top+'</td></tr>';});h+='</table>';h+='<div style="font-weight:700;color:var(--t1);margin-bottom:8px;font-family:inherit;font-size:12px">Detall complet <span style="color:var(--t3);font-weight:400">('+log.length+' registres'+(log.length>300?', mostrant 300':'')+' — IDs únics)</span></div>';h+='<div style="max-height:380px;overflow-y:auto;border:1px solid var(--br);border-radius:8px">';h+='<table style="width:100%;border-collapse:collapse;font-size:11px;font-family:monospace">';h+='<tr style="position:sticky;top:0;background:var(--bg3);color:var(--t3)"><td style="padding:4px 6px">T</td><td style="padding:4px 6px">Tipus act.</td><td style="padding:4px 6px">Usr</td><td style="padding:4px 6px">ID</td><td style="padding:4px 6px">Asumpte</td><td style="padding:4px 6px">Client</td></tr>';log.slice(0,300).forEach(function(x,i){var u=USERS[x.uid]||{ini:'?',color:'#888'};h+='<tr style="border-bottom:1px solid var(--br);background:'+(i%2?'transparent':'var(--bg3)')+'"><td style="padding:3px 6px;color:'+(x.t==='opo'?'#8B5CF6':'#3B82F6')+';font-weight:700">'+x.t.toUpperCase()+'</td><td style="padding:3px 6px;color:var(--t2)">'+x.ta+'</td><td style="padding:3px 6px;color:'+u.color+';font-weight:700">'+u.ini+'</td><td style="padding:3px 6px;color:var(--t4)">'+x.id+'</td><td style="padding:3px 6px;color:var(--t1)">'+esc(x.as).substring(0,45)+'</td><td style="padding:3px 6px;color:var(--t3)">'+esc(x.cl).substring(0,30)+'</td></tr>';});h+='</table></div>';box.innerHTML=h;var bC=document.createElement('button');bC.textContent='Tancar';bC.style.cssText='margin-top:16px;padding:7px 20px;background:#DC0028;color:#fff;border:none;border-radius:7px;cursor:pointer;font-family:inherit;font-size:12px;font-weight:600';bC.addEventListener('click',function(){ov.remove();});box.appendChild(bC);ov.appendChild(box);app.appendChild(ov);}
window._OAv={nd:function(d){if(S.view==='week')S.date.setDate(S.date.getDate()+d*7);else if(S.view==='month'){S.date.setMonth(S.date.getMonth()+d);S.date.setDate(1);}else S.date.setDate(S.date.getDate()+d);rHdr();rMC();if(S.all.length)applyF();},gt:function(){S.date=new Date();rHdr();rMC();if(S.all.length)applyF();},sv:function(v,el){S.view=v;qa('.vt').forEach(function(t){t.classList.remove('on');});el.classList.add('on');rHdr();if(S.all.length)applyF();},tu:function(uid,el){var tog=q('ut'+uid);if(S.aU.has(uid)){S.aU.delete(uid);tog.textContent='';tog.classList.remove('on');el.classList.remove('on');}else{S.aU.add(uid);tog.textContent='✓';tog.classList.add('on');el.classList.add('on');}applyF();},tt:function(t,el){if(S.aT.has(t)){S.aT.delete(t);el.style.opacity='.35';}else{S.aT.add(t);el.style.opacity='1';}applyF();},ta:function(){var all=Object.keys(TYPES).every(function(t){return S.aT.has(t);});if(all){S.aT.clear();}else{Object.keys(TYPES).forEach(function(t){S.aT.add(t);});}updTF();applyF();},ts:function(s,el){if(S.aS.has(s)){S.aS.delete(s);el.classList.add('soff');}
else{S.aS.add(s);el.classList.remove('soff');}
applyF();},ssf:function(k,el){if(S.sf===k||k==='all'){S.sf=null;qa('.si').forEach(function(x){x.classList.remove('on');});}else{S.sf=k;qa('.si').forEach(function(x){x.classList.remove('on');});el.classList.add('on');}applyF();},mc:function(d){mcOff+=d;rMC();},jd:function(y,m,d){S.date=new Date(y,m,d);mcOff=0;S.sf=null;qa('.si').forEach(function(x){x.classList.remove('on');});rHdr();rMC();if(S.view!=='day'){S.view='day';qa('.vt').forEach(function(t,i){t.classList[i===0?'add':'remove']('on');});}applyF();},tg:function(){var isDark=app.getAttribute('data-t')==='1';app.setAttribute('data-t',isDark?'0':'1');var tb=q('tb');if(tb)tb.textContent=isDark?'☀️':'🌙';},fa:async function(){var setCNT=function(msg){if(CNT)CNT.innerHTML='<div class="ld"><div class="sp"></div><span>'+msg+'</span></div>';};setCNT('Conectando con el CRM...');var opoCTX=null;try{var pageR=await fetch('/FE/ISC.Gaan.Oportunidad.FE.GridOportunidad/?_CIA=SCO');if(pageR.ok){var html=await pageR.text();var mm=html.match(/name="_GAAN_ClientContext_"[^>]*value='([^']+)'/);if(!mm)mm=html.match(/name="_GAAN_ClientContext_"[^>]*value="([^"]+)"/);if(mm){var cs=mm[1].replace(/&quot;/g,'"').replace(/&amp;/g,'&');opoCTX=JSON.parse(cs);}}}catch(e){}
if(!opoCTX)opoCTX=Object.assign({},CTX,{init:'ISC.Gaan.Oportunidad.FE.GridOportunidad'});var seenAct=new Set();var seenOpo=new Set();S.all=[];S.importLog=[];
var ACT_CFGS=[{targetUid:'M441819E',codAsoc:null,selMed:'M441819E'},{targetUid:'MA48168T',codAsoc:null,selMed:'MA48168T'},{targetUid:'M354046Y',codAsoc:null,selMed:'M354046Y'},{targetUid:'M441819E',codAsoc:'324781',selMed:'-1'},{targetUid:'MA48168T',codAsoc:'353621',selMed:'-1'},{targetUid:'M354046Y',codAsoc:'381799',selMed:'-1'}];
var OPO_CFGS=[{targetUid:'M441819E',codAsoc:null,selMed:'M441819E'},{targetUid:'MA48168T',codAsoc:null,selMed:'MA48168T'},{targetUid:'M354046Y',codAsoc:null,selMed:'M354046Y'},{targetUid:'M441819E',codAsoc:'324781',selMed:'-1'},{targetUid:'MA48168T',codAsoc:'353621',selMed:'-1'},{targetUid:'M354046Y',codAsoc:'381799',selMed:'-1'}];
var appendActs=function(dataArr,targetUid){dataArr.forEach(function(arr){arr.forEach(function(a){var uid=(a.UID_RACF||'').trim()||targetUid||'MA48168T';if(!USERS[uid])uid=targetUid||'MA48168T';if(seenAct.has(a.IDACTIV))return;seenAct.add(a.IDACTIV);a._uid=uid;S.importLog.push({t:'act',ta:gT(a),uid:uid,id:a.IDACTIV,as:a.DESASUN||'',cl:a.CLIENTE||''});S.all.push(a);});});};
var appendOpos=function(data,targetUid){data.forEach(function(o){var uid=(o.UID_RACF||'').trim()||targetUid||'MA48168T';if(!USERS[uid])uid=targetUid||'MA48168T';if(seenOpo.has(o.IDOPOACT))return;seenOpo.add(o.IDOPOACT);S.importLog.push({t:'opo',ta:'Oportunidad',uid:uid,id:o.IDOPOACT,as:o.ASUNTO||'',cl:o.CLIENTE||''});S.all.push({IDACTIV:o.IDOPOACT,IDOPOACT:o.IDOPOACT,_isOpo:true,_uid:uid,DESASUN:o.ASUNTO,CLIENTE:o.CLIENTE,TIM_INICI:o.TIMFREA||o.TIMULTI||o.TIMESTI||'',TIM_INICI_HORA:o.TIMFREA_HORA||'',TIM_FI:o.TIMESTI||'',TIM_FI_HORA:null,INDSITU:'0',INDSITU_DESC:o.INDSITU_DESC||'Abierta',INDACTI_DESC:'Oportunidad',INDPRIO_DESC:o.INDPRIORI_DESC,OBSACT:o.OBSACT,UID_RACF:uid,NOMUSER:o.NOMUSER,REFNUMPERS:o.REFNUMPERS,RAMEMIS_DESC:o.RAMEMIS_DESC,DATOPOR:o.DATOPOR,TIMULTI:o.TIMULTI,TIMFREA:o.TIMFREA,PRIEST:o.PRIEST,COMEST:o.COMEST,IDPROB:o.IDPROB});});};
var fetchActPage=async function(cfg,page,paginaCode,pila){var msg={SelectorMediador:{ClaveSeleccionada:cfg.selMed},SelectorVista:{ClaveSeleccionada:'208'},CheckMisActividades:{Valor:false},Concepto:{Valor:null},AE_COD_MEDIADOR_RACF:'00017167',CIAGRUPO:'SCO',ENTIDAD:'Appointment',OPERATIVA:'GridActividades',NAVEGACION:page===1?'CARGA_INICIAL':'CARGA_PAGINADO',COD_MEDIADOR_LOGADO:'-1',UsuarioRol:'Oficina'};if(cfg.codAsoc)msg.CodAsociados={ClaveSeleccionada:cfg.codAsoc};msg.ListaResultados={Paginacion:{TamanoPagina:500}};if(page>1&&paginaCode){msg.ListaResultados.Paginacion.PaginacionPila=pila.slice();msg.ListaResultados.Paginacion.PaginacionCodigoSiguiente=paginaCode;msg.ListaResultados.Paginacion.Paginando=true;}var ctx=Object.assign({},CTX,{init:'ISC.Gaan.Actividad.FE.GridActividad'});var body=new URLSearchParams({'_GAAN_Mensaje_Pantalla_':JSON.stringify(msg),'_GAAN_Comando_':page===1?'buscar':'paginar','_GAAN_ClientContext_':JSON.stringify(ctx),'_GAAN_SPA_':'true'});try{var r=await fetch('/FE/ISC.Gaan.Actividad.FE.GridActividad/',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:body.toString()});var j=await r.json();var lista=j&&j.message&&j.message.ListaResultados;var data=(lista&&lista.FuenteDatos)||[];var pag=lista&&lista.Paginacion;return{data:data,nextCode:pag&&pag.PaginacionCodigoSiguiente||null};}catch(e){return{data:[],nextCode:null};}};
var fetchOpoPage=async function(cfg,page,paginaCode,pila){var msg={SelectorMediador:{ClaveSeleccionada:cfg.selMed},SelectorVista:{ClaveSeleccionada:'146'},CheckMisOportunidades:{Valor:false},AE_COD_MEDIADOR_RACF:'00017167',CIAGRUPO:'SCO',ENTIDAD:'Oportunidad',OPERATIVA:'GridOportunidades',NAVEGACION:page===1?'CARGA_INICIAL':'CARGA_PAGINADO',COD_MEDIADOR_LOGADO:'-1',UsuarioRol:'Oficina'};if(cfg.codAsoc)msg.CodAsociados={ClaveSeleccionada:cfg.codAsoc};msg.ListaResultados={Paginacion:{TamanoPagina:500}};if(page>1&&paginaCode){msg.ListaResultados.Paginacion.PaginacionPila=pila.slice();msg.ListaResultados.Paginacion.PaginacionCodigoSiguiente=paginaCode;msg.ListaResultados.Paginacion.Paginando=true;}var ctx=Object.assign({},opoCTX,{init:'ISC.Gaan.Oportunidad.FE.GridOportunidad'});var body=new URLSearchParams({'_GAAN_Mensaje_Pantalla_':JSON.stringify(msg),'_GAAN_Comando_':page===1?'buscar':'paginar','_GAAN_ClientContext_':JSON.stringify(ctx),'_GAAN_SPA_':'true'});try{var r=await fetch('/FE/ISC.Gaan.Oportunidad.FE.GridOportunidad/',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:body.toString()});var j=await r.json();var lista=j&&j.message&&j.message.ListaResultados;var data=(lista&&lista.FuenteDatos)||[];var pag=lista&&lista.Paginacion;return{data:data,nextCode:pag&&pag.PaginacionCodigoSiguiente||null};}catch(e){return{data:[],nextCode:null};}};
var MAX_PAGES=40;var loadAllActs=async function(){for(var i=0;i<ACT_CFGS.length;i++){var cfg=ACT_CFGS[i],code=null,pila=[];for(var p=1;p<=MAX_PAGES;p++){var iCode=USERS[cfg.targetUid]?USERS[cfg.targetUid].ini:cfg.targetUid;setCNT('Cargando ACTS ('+iCode+')… (pág '+p+')');var r=await fetchActPage(cfg,p,code,pila);if(r.data.length)appendActs([r.data],cfg.targetUid);if(!r.nextCode)break;pila.push(code||'1');code=r.nextCode;}}};
var loadAllOpos=async function(){for(var i=0;i<OPO_CFGS.length;i++){var cfg=OPO_CFGS[i],code=null,pila=[];for(var p=1;p<=MAX_PAGES;p++){var iCode=USERS[cfg.targetUid]?USERS[cfg.targetUid].ini:cfg.targetUid;setCNT('Cargando OPOS ('+iCode+')… (pág '+p+')');var r=await fetchOpoPage(cfg,p,code,pila);if(r.data.length)appendOpos(r.data,cfg.targetUid);if(!r.nextCode)break;pila.push(code||'1');code=r.nextCode;}}};
await loadAllActs();await loadAllOpos();var asg=loadAsg();S.all.forEach(function(a){if(asg[a.IDACTIV]&&USERS[asg[a.IDACTIV]]){a._uid=asg[a.IDACTIV];}});
        // Carregar overrides de OneDrive i aplicar-los
        setCNT('Carregant overrides OneDrive...');
        await _loadOverrides();
        _applyOverrides();
        setCNT('Procesando…');applyF();_updHeader();
        setTimeout(function(){
          var hr=SR.querySelector('.hr');
          if(hr&&!SR.getElementById('impRptBtn')){
            var rb=document.createElement('button');rb.id='impRptBtn';rb.className='thb';rb.title='Informe d\'importació';rb.textContent='📊';rb.addEventListener('click',function(){showImportRpt();});hr.insertBefore(rb,hr.firstChild);
          }
          // Botó OneDrive: verd si connectat, groc si cal login
          if(hr&&!SR.getElementById('odBtn')){
            var odBtn=document.createElement('button');odBtn.id='odBtn';odBtn.className='rfb';
            var connected=OD.isReady();
            odBtn.textContent=connected?'☁️ OD':'🔑 OD';
            odBtn.title=connected?'OneDrive connectat — overrides actius':'Connectar OneDrive per overrides multi-usuari';
            odBtn.style.cssText=connected?'background:rgba(16,185,129,.12);border-color:rgba(16,185,129,.35);color:#10B981':'background:rgba(234,179,8,.12);border-color:rgba(234,179,8,.35);color:#CA8A04';
            odBtn.addEventListener('click',async function(){
              if(OD.isReady()){alert('OneDrive ja connectat. '+Object.keys(_overrides).length+' overrides actius.');return;}
              odBtn.textContent='...';
              var ok=await OD.loginPopup();
              if(ok){
                await _loadOverrides();_applyOverrides();applyF();
                odBtn.textContent='☁️ OD';odBtn.style.cssText='background:rgba(16,185,129,.12);border-color:rgba(16,185,129,.35);color:#10B981';
                odBtn.title='OneDrive connectat — '+Object.keys(_overrides).length+' overrides actius';
              } else {
                odBtn.textContent='🔑 OD';alert('Login OneDrive cancel·lat o fallat.');
              }
            });
            hr.insertBefore(odBtn,hr.firstChild);
          }
        },300);},ca:async function(id){var a=S.all.find(function(x){return x.IDACTIV===id;});if(!a)return;if(!confirm('¿Completar actividad?'))return;var prev=a.INDSITU_DESC;a.INDSITU_DESC='Completada';applyF();var ctx=Object.assign({},CTX,{init:'ISC.Gaan.Actividad.FE.DetalleActividad'});var msg={IDACTIV:id,UID_RACF:a._uid||a.UID_RACF||'MA48168T',AE_COD_MEDIADOR_RACF:'00017167',CIAGRUPO:'SCO'};var body=new URLSearchParams({'_GAAN_Mensaje_Pantalla_':JSON.stringify(msg),'_GAAN_Comando_':'completar','_GAAN_ClientContext_':JSON.stringify(ctx),'_GAAN_SPA_':'true'});try{var r=await fetch('/FE/ISC.Gaan.Actividad.FE.DetalleActividad/',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:body.toString()});if(!r.ok)throw new Error('fallo req');}catch(e){a.INDSITU_DESC=prev;applyF();alert('No se pudo completar en el CRM.');}},oa:function(id){window.openCRM('ISC.Gaan.Actividad.FE.DetalleActividad','IDACTIV',id);},oc:function(ref){window.openCRM('PER.Clientes.V360.FE.V360','IdReferencia',ref);},diag:async function(){var btn=q('diagb');if(btn){btn.textContent='...';}
var weekDates=gWk(new Date()).map(function(d){return fmtD(d);});var doF=async function(sel,vista,chk){var msg={SelectorMediador:{ClaveSeleccionada:sel},SelectorVista:{ClaveSeleccionada:vista},Concepto:{Valor:null},CheckMisActividades:{Valor:!!chk},AE_COD_MEDIADOR_RACF:'00017167',CIAGRUPO:'SCO',UID_RACF:sel,ENTIDAD:'Appointment',OPERATIVA:'GridActividades',NAVEGACION:'CARGA_INICIAL',COD_MEDIADOR_LOGADO:'-1',UsuarioRol:'Oficina'};var ctx=Object.assign({},CTX,{init:'ISC.Gaan.Actividad.FE.GridActividad'});var body=new URLSearchParams({'_GAAN_Mensaje_Pantalla_':JSON.stringify(msg),'_GAAN_Comando_':'buscar','_GAAN_ClientContext_':JSON.stringify(ctx),'_GAAN_SPA_':'true'});try{var ac=new AbortController();setTimeout(function(){ac.abort();},8000);var r=await fetch('/FE/ISC.Gaan.Actividad.FE.GridActividad/',{method:'POST',signal:ac.signal,headers:{'Content-Type':'application/x-www-form-urlencoded'},body:body.toString()});var j=await r.json();var data=(j&&j.message&&j.message.ListaResultados&&j.message.ListaResultados.FuenteDatos)||[];var wc=data.filter(function(a){return weekDates.indexOf(a.TIM_INICI||'')>=0;}).length;return'total='+data.length+' semana='+wc;}catch(e){return'ERR';}};var tests=[['M354046Y','115',0],['M354046Y','208',0],['M441819E','115',0],['MA48168T','115',0],['TODOS','115',0],['TODOS','208',0],['M354046Y','115',1],['M354046Y','208',1]];var labels=['FK v115','FK v208','DH v115','SF v115','TODOS v115','TODOS v208','FK v115 chkT','FK v208 chkT'];var results=await Promise.all(tests.map(function(t){return doF(t[0],t[1],t[2]);}));var ov=document.createElement('div');ov.style.cssText='position:absolute;inset:0;background:rgba(0,0,0,.82);display:flex;align-items:center;justify-content:center;z-index:10;';var box=document.createElement('div');box.style.cssText='background:var(--bg2);border:2px solid var(--brl);border-radius:12px;padding:22px;max-width:480px;width:100%;font-size:12px;line-height:2;font-family:monospace';var hdr=document.createElement('div');hdr.style.cssText='font-weight:700;font-size:14px;margin-bottom:10px;color:var(--t1);font-family:inherit';hdr.textContent='Diagnostico API — resultados';box.appendChild(hdr);results.forEach(function(res,i){var d=document.createElement('div');var n=Number((res.match(/total=(\d+)/)||[0,0])[1]);d.style.color=n>100?'#10B981':n>50?'#F59E0B':'var(--t3)';d.textContent=labels[i]+': '+res;box.appendChild(d);});var bC=document.createElement('button');bC.textContent='Cerrar';bC.style.cssText='margin-top:14px;padding:6px 16px;background:#DC0028;color:#fff;border:none;border-radius:6px;cursor:pointer;font-family:inherit;font-size:12px';bC.addEventListener('click',function(){ov.remove();if(btn)btn.textContent='Diag';});box.appendChild(bC);ov.appendChild(box);app.appendChild(ov);},tRt:function(rt){if(S.rts.has(rt)){if(S.rts.size>1)S.rts.delete(rt);}
else S.rts.add(rt);['act','opo'].forEach(function(k){var el=SR.getElementById('rtf-'+k);if(!el)return;el.className='rtfb '+k+(S.rts.has(k)?' on':'');});if(S.all.length)applyF();},close:function(){document.getElementById(HOST)?.remove();},srch:function(qtxt){S.sq=(qtxt||'').trim().toLowerCase();var clr=SR.getElementById('srch-clr');if(clr)clr.className='srch-clr'+(S.sq?' vis':'');if(S.sq.length>0&&S.view!=='kanban'){var kBtn=SR.getElementById('tab-kanban');_OAv.sv('kanban',kBtn);return;}if(S.all.length)applyF();},srchSt:function(v){},da:async function(id){var a=S.all.find(function(x){return x.IDACTIV===id;});if(!a)return;if(!confirm('\xbfCancelar esta actividad en el CRM?'))return;var prev=a.INDSITU_DESC;a.INDSITU_DESC='Cancelada';applyF();var ok=await crmPost('ISC.Gaan.Actividad.FE.DetalleActividad','cancelar',{IDACTIV:id,UID_RACF:a._uid||'M441819E',AE_COD_MEDIADOR_RACF:'00017167',CIAGRUPO:'SCO'});if(!ok){a.INDSITU_DESC=prev;applyF();alert('No se pudo cancelar en el CRM.');}
else{S.all=S.all.filter(function(x){return x.IDACTIV!==id;});applyF();}},rua:function(actId,uid){var a=S.all.find(function(x){return x.IDACTIV===actId;});if(!a)return;if(uid&&USERS[uid]){a._uid=uid;var asg=loadAsg();asg[actId]=uid;saveAsg(asg);}updC();render();},vtosToggle:function(){S.vtos.show=!S.vtos.show;var btn=q('veye');if(btn)btn.classList.toggle('on',S.vtos.show);if(S.view!=='day'){S.view='day';qa('.vt').forEach(function(t,j){t.classList[j===0?'add':'remove']('on');});rHdr();}render();},vtosNew:function(){vtosModal(null);},vtosEdit:function(id){var it=S.vtos.items.find(function(x){return x.id===id;});if(it)vtosModal(it);},vtosSave:function(id){var asunto=(q('va')?.value||'').trim();if(!asunto){alert('El asunto es obligatorio');return;}var it={id:id||Date.now()+'',asunto:asunto,cliente:(q('vc')?.value||'').trim(),fecha:(q('vfe')?.value||fmtD(S.date)).trim(),tipo:q('vti')?.value||'Tarea',notas:(q('vn')?.value||'').trim(),creado:Date.now()};if(id){S.vtos.items=S.vtos.items.map(function(x){return x.id===id?it:x;});}else{S.vtos.items.unshift(it);}saveVtos(S.vtos.items);q('vmodal')?.remove();updVtos();if(S.view==='day')render();},vtosDelete:function(id){if(!confirm('¿Eliminar esta entrada?'))return;S.vtos.items=S.vtos.items.filter(function(x){return x.id!==id;});saveVtos(S.vtos.items);updVtos();if(S.view==='day')render();},vtosAssign:function(vid,uid){var it=S.vtos.items.find(function(x){return x.id===vid;});if(!it)return;var u=USERS[uid];if(!confirm('¿Asignar "'+it.asunto+'" a '+u.name+'?\nSe abrirá el CRM para crear la actividad.'))return;window.open('/FE/ISC.Gaan.Actividad.FE.DetalleActividad/?_CIA=SCO','_blank');S.vtos.items=S.vtos.items.filter(function(x){return x.id!==vid;});saveVtos(S.vtos.items);updVtos();render();},showDetail:showDetail};document.addEventListener('keydown',function(e){if(e.key==='Escape'){if(q('vmodal')){q('vmodal').remove();return;}document.getElementById(HOST)?.remove();}});rHdr();rMC();updTF();updVtos();_OAv.fa();})();