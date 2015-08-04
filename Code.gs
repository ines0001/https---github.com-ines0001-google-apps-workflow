function doGet(request) {
  
  var html=null, duration;
  
  
  switch(request.parameter.page){
  
    case 'new':html = HtmlService.createTemplateFromFile('Page-new')
               duration = PropertiesService.getScriptProperties().getProperty('ID_DURATION') 
               html.duration = (duration)?duration:'40s';
               
               html = html.evaluate()
                        .setTitle('Saisie formulaire AO')
                        .setSandboxMode(HtmlService.SandboxMode.IFRAME);break;
    
    case 'dashboard':html = HtmlService.createTemplateFromFile('Page-dashboard');
                         html.clsid = CLSID;  
                        html=html.evaluate()
                        .setTitle('Administration AO')
                        .setSandboxMode(HtmlService.SandboxMode.IFRAME);break;                    
                        
                        
    case 'response':html = HtmlService.createTemplateFromFile('Page-response');
                    html.reference = request.parameter.ref;
                    html.type= (request.parameter.type=='expanded')?request.parameter.type:'small';
                    html = html.evaluate()
                        .setTitle('Réponse formulaire AO')
                        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
                    
                        
                    break;
      
    case 'relance':html = HtmlService.createTemplateFromFile('Page-relance');
                    html.reference = request.parameter.ref; 
                    html = html.evaluate()
                               .setTitle('Relance formulaire AO')
                               .setSandboxMode(HtmlService.SandboxMode.IFRAME);
      
      
                    break;
      
    case 'decision':html = HtmlService.createTemplateFromFile('Page-decision');
                    html.result = request.parameter.result;              
                    html.reference = request.parameter.ref;
                    html = html.evaluate()
                               .setTitle('Decision formulaire AO')
                               .setSandboxMode(HtmlService.SandboxMode.IFRAME);
      
      
                    break;
    
    
    case 'testing':html = HtmlService.createTemplateFromFile('testing');
                    html.reference = request.parameter.ref;
                    html.duration = '40s';
                    
                    
                    html.ticker = enumParams('Informations flash')
                    
                    html = html.evaluate()
                               .setTitle('Testing formulaire AO')
                               .setSandboxMode(HtmlService.SandboxMode.IFRAME);
      
      
                    break;
      
      
    case 'build_avv':html = HtmlService.createTemplateFromFile('Page-build_avv');
                    html.reference = request.parameter.ref;
                    html = html.evaluate()
                               .setTitle('Building AVV')
                               .setSandboxMode(HtmlService.SandboxMode.IFRAME);
      
      
                    break;
                    
    case 'admin': /* Necessite droit en écriture sur le script
                  if(request.parameter.debugging){
                          Logger.setLevel(request.parameter.debugging);
                          PropertiesService.getScriptProperties().setProperty('ID_LOG_LEVEL',Logger.getLevel());
                          Logger.severe('level debugging is now:%s',Logger.getLevel());
                  }
                  */
                  var clsid_ = PropertiesService.getScriptProperties().getProperty('ID_CLSID')
                  /* Necessite droit en écriture sur le script
                  if(request.parameter.clsid){
                          clsid_ = (request.parameter.clsid.toUpperCase()=='PROD')?ID_CLSID_PROD:ID_CLSID_DEV
                          PropertiesService.getScriptProperties().setProperty('ID_CLSID',clsid_)
                          Logger.severe('clsid is now:%s',clsid_);
                  }
                  */
                  
                  var out = '<h1>SQLI PP9-WEB</h1><b><p>Command : Admin</p></b><p><i>debugging:</i><b>'+Logger.getLevel()+'</p></b></br>';
                  out += '<p><i>clsid_dev (spreadsheet):</i><b>'+clsid_+'</p></b>';
                  out += '<p><i>production:</i><b>'+__PROD__+'</p></b>';
                  out += '<p><i>url:</i><b>'+URL_EXEC+'</p></b>';
                  
                  
              
                  html = HtmlService.createHtmlOutput(out); 
                    
                  break;                    
                        
    default:html = HtmlService.createTemplateFromFile('Page-accueil')
            html.duration = (duration)?duration:'40s';
            html = html.evaluate()
                          .setTitle('Accueil formulaire AO')
                          .setSandboxMode(HtmlService.SandboxMode.IFRAME);break;
 }
 
 return html;
  
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

// BEGIN DRIVE INTERFACE
/* fonction permettant de déplacer les fichiers(files) vers
   le dossier drive (destination). La fonction permet de créer
   le folder destination si celui-ci n'existe pas. Tous les fichiers
   sources (IDs) seront ensuite effacés de la source DRIVE du user
   files : Array ID drive file
   destination: string define the name folder destination
*/
function saveFiles(files,destination,subdestination){
try{
    var file, folder = DriveApp.getFolderById(ID_FOLDER_FILES);
    
    
    var folder_dest = folder.getFoldersByName(destination);
    if(__DEBUG__) {Logger.finest(JSON.stringify(files));}
    /* verification de l'existence du dossier destination */
    if(!folder_dest.hasNext()){
        folder_dest = folder.createFolder(destination);
    }else folder_dest = folder_dest.next();
    
    /* Existance d'un sous dossier */
    var sub_folder;
    
  if(subdestination===undefined){
    sub_folder = '';
  }else{
    sub_folder= folder_dest.getFoldersByName(subdestination);
    if(!sub_folder.hasNext()){
        sub_folder = folder_dest.createFolder(subdestination);
    }else sub_folder = sub_folder.next();
    folder_dest = sub_folder;
  }
    
  
  
    /* recopie de l'ensemble des fichiers dans le folder destination*/
    
    for(var index=0; index< files.length;index++){
      if(files[index]===undefined || typeof files[index]!='string' || files[index]=='' ) continue;
      
      file = DriveApp.getFileById(files[index]);
      file.makeCopy(folder_dest);
      
      DriveApp.removeFile(file);
    }
 
    return folder_dest.getUrl();
}catch(e){treatmentException_(e)}   
}

function testing_saveFiles(){
 try{ 
  var r = saveFiles(['0BxTfS7jXR_FYZThDRG5GT1B2Z2s'],'I1429787869176','toto');
}catch(e){treatmentException_(e)}   
  
  
}

function getInfoGeneric(ref){
 try{ 
  
  var objet = new SHEET_AO();
  var row = objet.Info_A(ref);
  var row_b = objet.Info_B(ref);
  
  Logger.log('getInfoGeneric: %s',JSON.stringify(objet,null,null));
  
  if( row&&row_b) 
          return {sale:row.emetteur,
                          client:row.client,
                          dossier:row.dossier,
                          remise:Utilities.formatDate(row_b.remise, "GMT+02:00", "dd/MM/yyyy"),
                          url:row.url,
                          manager:row.managerEnCharge,
                          rp:row.rp,
                          dar:(row.dateDarPlanifie===undefined )?'':Utilities.formatDate(row.dateDarPlanifie, "GMT+02:00", "dd/MM/yyyy"),
                          visa:(row.dateVisaForfaitPrvu===undefined )?'':Utilities.formatDate(row.dateVisaForfaitPrvu, "GMT+02:00", "dd/MM/yyyy"),
                          state:row.statut,
                          historical:row.log,};
  else throw new Error('Aucun dossier associé avec la référence suivante : '+ref);
 }catch(e){treatmentException_(e)} 
}

function testing(){

  try{
    
    // 0BxTfS7jXR_FYSDRFQ1Z2OXk4cEk
    var fileID = '0BxTfS7jXR_FYSDRFQ1Z2OXk4cEk';

    saveFiles(fileID.split(","),'testing','CR_GONOGO');
    
 
    
  }catch(e){Logger.log(e);}

}

// END DRIVE INTERFACE






function getHtml_practice(){
 try{ 
  var count=0,html='<option>'+INIT_PRACTICE_VALUE+'</option>';
  
  var found=true;
  var column ={Unit:-1,Practice:-1};
  var params = SpreadsheetApp.openById(CLSID)
                         .getSheetByName(SHEET_PARAMS)
                         .getDataRange()
                         .getValues();
  
  
  
  for( var i=0;i< params[0].length;i++){
    if(params[0][i]==COLUMN_PRACTICE ) {column.Practice=i;}
    if(params[0][i]==COLUMN_UNIT ) {column.Unit=i;}
    
  }
  
  for( var i in column){ if(column[i]==-1) found=false;}
  
    
  if(column.Practice==-1) {throw "no params";}
  
  params = SpreadsheetApp.openById(CLSID)
                         .getSheetByName(SHEET_PARAMS)
                         .getRange(2,column.Practice+1,50)
                         .getValues();
  
  
  var out = cleanArray(params);
  
  
  for( var i=0;i<out.length;i++){
    html+='<option>'+out[i]+'</option>';
    
    
  }
  
  return html;
 }catch(e){treatmentException_(e)}    
}


function getHtml_unit(){
 try{
  var count=0,html='';
  
  var found=true;
  var column ={Unit:-1,Practice:-1,Unit_press:-1};
  var params = SpreadsheetApp.openById(CLSID)
                         .getSheetByName(SHEET_PARAMS)
                         .getDataRange()
                         .getValues();
  
    
  for( var i=0;i< params[0].length;i++){
    if(params[0][i]==COLUMN_PRACTICE ) {column.Practice=i;}
    if(params[0][i]==COLUMN_UNIT ) {column.Unit=i;}
    if(params[0][i]==COLUMN_UNIT_PRESS ) {column.Unit_press=i;}
    
  }
  
  if(column.Practice==-1 && column.Unit==-1 && column.Unit_press==-1) {throw "no params";}
  
    
  for( var i=1;i<params.length;i++)
    
     if(params[i][column.Unit_press]) html+='<option>'+params[i][column.Unit_press]+'</option>';
 
 
  
  return html;
 }catch(e){treatmentException_(e)}    
}

function GetUnit(pratice){
 try{ 
  var unit='', found=true;
  
  var column ={Unit:-1,Practice:-1,Unit_press:-1};
  var params = SpreadsheetApp.openById(CLSID)
                         .getSheetByName(SHEET_PARAMS)
                         .getDataRange()
                         .getValues();
  
  for( var i=0;i< params[0].length;i++){
    if(params[0][i]==COLUMN_PRACTICE ) {column.Practice=i;}
    if(params[0][i]==COLUMN_UNIT ) {column.Unit=i;}
    if(params[0][i]==COLUMN_UNIT_PRESS ) {column.Unit_press=i;}
    
  }
  
  for( var i in column){ if(column[i]==-1) found=false;}
  
    
  if(column.Practice==-1 && column.Unit==-1 && column.Unit_press==-1) {throw "no params";}
  
    
  for( var i=0;i<params.length;i++){
    
    if( pratice!=INIT_PRACTICE_VALUE && params[i][column.Practice]==pratice) { unit = params[i][column.Unit];}
 
    
    
  }
  
  
  return unit;
  
  }catch(e){treatmentException_(e)} 
}


function HtmlPreviewModal(form,mail){
 try{ 
  if(form===undefined || typeof form!='object') throw 'null form';
  
  if(__DEBUG__) Logger.finest('HtmlPreviewModal:'+mail);
   return MailingToManager(Session.getActiveUser().getEmail(),
                   '#',
                   form,
                   mail,
                   false);
  }catch(e){treatmentException_(e)} 
}

function GetManagers(){
  var column ={Unit:-1,Practice:-1,Unit_press:-1}, out=[];
  var params = SpreadsheetApp.openById(CLSID)
                         .getSheetByName(SHEET_PARAMS)
                         .getDataRange()
                         .getValues();
  
  for( var i=0;i< params[0].length;i++){
    if(params[0][i]==COLUMN_UNIT_PRESS ) {column.Unit_press=i;break;}
       
  } 

  if(column.Unit_press==-1 ) {throw "no params";}
  for( var i=1;i<params.length;i++){
    if(params[i][column.Unit_press]!='') out.push(new Array(params[i][column.Unit_press],params[i][column.Unit_press+1]));
    
  }
  return out;  
}

function GetManager(unit){
  try{
  if(unit===undefined ) throw 'null unit';
  
  var column ={Unit:-1,Practice:-1,Unit_press:-1}, out='';
  var params = SpreadsheetApp.openById(CLSID)
                         .getSheetByName(SHEET_PARAMS)
                         .getDataRange()
                         .getValues();
  
  for( var i=0;i< params[0].length;i++){
    if(params[0][i]==COLUMN_UNIT_PRESS ) {column.Unit_press=i;break;}
       
  } 

  if(column.Unit_press==-1 ) {throw "no params";}
  
  for( var i=0;i<params.length;i++){
    

    if( params[i][column.Unit_press]==unit) { out = params[i][column.Unit_press+1]; }
    
  }
  if(__DEBUG__) Logger.finest('GetManager: %s',JSON.stringify(out,null,'\t'));


  return out;
  }catch(e){treatmentException_(e)} 
}

function save_form(form,mail,filesID){
 try{
  var url_drive=null,out = [],out2 =[],ref='I',objet_log= new OLogger();
 
  if(form===undefined || typeof form!='object') throw 'null form';
  
  Logger.log('save_form %s',JSON.stringify(form,null,null));
 
  var date = new Date(), state=GetRowParams(COLUMN_STATE);

  //Calcul de la référence dossier
  ref+=String(date.valueOf());
  // url googledrive du folder de stockage
  if( filesID) url_drive = saveFiles(filesID.split(","),ref);
  
  if( url_drive ) 
      mail+='<hr>Accès aux documents de l\'appel d\'offre: dossier Google Drive <a href="'
        +url_drive
        +'" target="_blank" title="Lien dossier google drive" >'+ref+'</a><hr>';
  else
      mail+='<hr>Aucun document associé<hr>';
  
  out.push(ref);
  out.push(Session.getActiveUser().getEmail());
  out.push(date);
  out.push(form.find('client'));
  out.push(form.find('ao'));
  out.push(form.find('practice'));
  out.push(GetManager(form.find('unit')));
  //\/\/\/\/\/\/BEGIN MANTIS 011
  out.push(form.find('date'),'','','','','','');
  //\/\/\/\/\/\/END MANTIS 011
  out.push(state.enAttente);
  out.push('','','','','','','','','',url_drive,objet_log.pushSession('Création dossier',state.enAttente, {client:form.find('client'),
                                                                                                           dossier:form.find('ao'),
                                                                                                           /* MANTIS 050 */
                                                                                                           manager_pressenti:GetManager(form.find('unit')),
                                                                                                           remise:form.find('date'),
                                                                                                           sales:Session.getActiveUser().getEmail(),}));
  out2.push(ref);
  out2.push(form.find('contexte'), form.find('enjeux'), form.find('unit'), form.find('budget'), form.find('date'), form.find('indice'), 
           form.find('critere'), form.find('origine'), form.find('pourquoi'), form.find('partenaires'),form.find('concurrents'), form.find('crm').toUpperCase(),
           url_drive,mail,form.find('techno')+'-'+form.find('techno-autre'),null);
  
  
  SpreadsheetApp.openById(CLSID)
                         .getSheetByName(SHEET_O_AVV)
                         .appendRow(out);
  SpreadsheetApp.openById(CLSID)
                         .getSheetByName(SHEET_O_HISTO)
                         .appendRow(out2);

  /*\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\\/\/\/\/\/\/\/\/\/\/
  
  Recopie la mise en forme complète d'une nouvelle ligne
  (lastRow,8, 1, 20) de SHEET_O_MASQUE_ROW vers SHEET_O_AVV
  \/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\\/\/\/\/\/\/\/\*/
  var lastRow = SpreadsheetApp.openById(CLSID)
                         .getSheetByName(SHEET_O_AVV)
                         .getLastRow()
  
  var range_dest = SpreadsheetApp.openById(CLSID)
                         .getSheetByName(SHEET_O_AVV)
                         .getRange(lastRow,9, 1, 16);
  
  var range_source = SpreadsheetApp.openById(CLSID)
                         .getSheetByName(SHEET_O_MASQUE_ROW)
                         .getRange(2, 9, 1, 16);
  
  
   
  range_source.copyTo(range_dest);
  SpreadsheetApp.flush();
  
  /*\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
  
  Diffusion du mail aux différents destinataires
  \/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\*/
  
  MailingToManager(Session.getActiveUser().getEmail(),
                   ref,
                   form,
                   mail,
                   true);
  
 return {id:ref,manager:GetManager(form.find('unit')),unit:form.find('unit'),drive:url_drive,};
 
 }catch(e){treatmentException_(e)} 
 
}

// function : onValidateRedirectionUnit
// Params : form.selectunit, form.commentaire_manager,form.identifiant
// Traitment : Envoyer un mail de notification au nouveau manager afin de prendre en charge le dossier
// return : à définir
function onValidateRedirectionUnit(form){
  try{
    if(__DEBUG__) Logger.finest('onValidateRedirectionUnit: form : %s ',JSON.stringify(form,null,'\t'));
    if(form===undefined ) throw 'null form';
    if(form.selectunit=='' ) throw 'null form';
    if(form.identifiant=='' ) throw 'null form';
    
    var ret=null,sheet =  SpreadsheetApp.openById(CLSID)
                           .getSheetByName(SHEET_O_AVV),
        sheet_form = SpreadsheetApp.openById(CLSID)
                           .getSheetByName(SHEET_O_HISTO); 
      
    var state=GetRowParams(COLUMN_STATE);
    
      
    // ****** FUNCTION INLINE  ****
    var row = (function () {
      var row, data =getRowsData(sheet, sheet.getRange(2, 1, sheet.getMaxRows() - 1, MAX_COL));
      
      
      for(var i=0;i<data.length;i++){
        row = data[i];
        row.rowNumber = i+2;
        if(form.identifiant==row.identifiant && row.statut!=state.enCours) return row;
        
      }
      return null;
      })();
    
    var row_sheet = (function () {
      var row, data =getRowsData(sheet_form, sheet_form.getRange(2, 1, sheet_form.getMaxRows() - 1, MAX_COL));
      
      
      for(var i=0;i<data.length;i++){
        row = data[i];
        row.rowNumber = i+2;
        if(form.identifiant==row.identifiant && row.statut!=state.enCours) return row;
        
      }
      return null;
      })();
    
    // ****** FUNCTION INLINE  ****
    
    if(row&&row_sheet) {
      
                  
                  sheet.getRange(row.rowNumber,COLUMN_SHEET_RU).setValue(form.selectunit); /*\/\/\/\/\/\/On affecte le nouveau Manager en suivi */
                  //sheet.getRange(row.rowNumber,COLUMN_SHEET_MANAGER).setValue(GetManager(form.selectunit));
                  sheet.getRange(row.rowNumber,COLUMN_SHEET_STATE).setValue(state.relance);
                  sheet.getRange(row.rowNumber,COLUMN_SHEET_DATE_RESP).setValue(new Date());
                  var objet_log= new OLogger(row.log);
                  sheet.getRange(row.rowNumber,COLUMN_SHEET_LOG).setValue(objet_log.pushSession('Dossier redirigé',state.relance,{client:row.client,
                                                                                                                                  dossier:row.dossier,
                                                                                                                                  info:form}));
                  //REnvoie du mail au Manager de UNIT
          
                  MailingToOtherManager(row.emetteur,
                                       {client:row.client,
                                        ao:row.dossier,
                                        contexte:row_sheet.contexte,
                                        enjeux:row_sheet.enjeux,
                                        unit:row_sheet.unit,
                                        budget:row_sheet.budget,
                                        date:row_sheet.remise,
                                        indice:row_sheet.indice,
                                        critere:row_sheet.critere,
                                        origine:row_sheet.origine,
                                        pourquoi:row_sheet.pourquoi,
                                        partenaires:row_sheet.partenaires,
                                        concurrents:row_sheet.concurrents,
                                        crm:row_sheet.crm,
                                        emetteur:row.emetteur,
                                        identifiant:row.identifiant,
                                        bodymail:row_sheet.message,
                                        url:row_sheet.url,
                                        manager:form.selectunit,
                                        techno:row_sheet.techno,

                                        previous_manager:Session.getActiveUser().getEmail(),
                                        info_complementaire: form.commentaire_manager,
                                                      
                                       });
                  
                 
                  ret=form.selectunit;
        
    }
   
    if(ret==null) throw 'le dossier r&eacute;f&eacute;rence <b>'+form.identifiant+'</b> est incomplet ou d&eacute;j&agrave; affect&eacute; au responsable de proposition <br>A bientôt,'
    return ret;
   }catch(e){treatmentException_(e)} 
}

// function : onValidateForRp
// Params : form.selectrp, form.commentaire_sale, form.date, form.identifiant
// Traitment : Envoyer un mail de notification au RP afin qu'il prenne en charge le dossier+envoie d'une notification google calendar
// return : à définir
function onValidateForRp(form){
  try{
    if(form===undefined ) throw 'null form';
    if(form.selectrp=='' ) throw 'null form';
    if(form.identifiant=='' ) throw 'null form';
    
    if(__DEBUG__) { Logger.finest('onValidateForRp : %s',JSON.stringify(form,null,null));}
    
    var ret=null, sheet =  SpreadsheetApp.openById(CLSID)
                           .getSheetByName(SHEET_O_AVV);
      
    var lastRow = sheet.getLastRow(), state=GetRowParams(COLUMN_STATE);
    
    var objet = new SHEET_AO();
    var row = objet.Info_A(form.identifiant);
    var row_b = objet.Info_B(form.identifiant);
    /*
      BUGFIX : Erreur sur call MailingToRP avec manager:row.managerEnCharge. Si pas de réaffectation alors le champ est vide et 
      cela génère une exception sur l'envoie du mail (notification.gs
    */
    var manager = row.managerEnCharge; 
    
    if( row&&row_b&&(row.statut==state.enAttente||row.statut==state.relance))
    {
      
      sheet.getRange(row.rowNumber,COLUMN_SHEET_RP).setValue(form.selectrp);
      /*\/\/\/\/\/\/\/\/\/\/\/\/
      BEGIN:MANTIS 0019
      \/\/\/\/\/\/\/\/\/\/*/
      // Vérification qu'une reaffectation de UNIT n'a pas été traité auparavant
      if(sheet.getRange(row.rowNumber,COLUMN_SHEET_RU).isBlank()){
        sheet.getRange(row.rowNumber,COLUMN_SHEET_MANAGER).copyTo(sheet.getRange(row.rowNumber,COLUMN_SHEET_RU));
        manager = row.managerPressenti
        
      }
      /*\/\/\/\/\/\/\/\/\/\/\/\/
      END:MANTIS 0019
      \/\/\/\/\/\/\/\/\/\/*/
      sheet.getRange(row.rowNumber,COLUMN_SHEET_DATE_GONOGO).setValue(form.date);
      sheet.getRange(row.rowNumber,COLUMN_SHEET_STATE).setValue(state.enCours);
      sheet.getRange(row.rowNumber,COLUMN_SHEET_DATE_RESP).setValue(new Date());
      
      var objet_log= new OLogger(row.log);
      sheet.getRange(row.rowNumber,COLUMN_SHEET_LOG).setValue(objet_log.pushSession('Dossier affecté',state.enCours,{client:row.client,
                                                                                                                     dossier:row.dossier,
                                                                                                                     info:form}));
                  
      
      // Push Event Calendar
      var calendar = CalendarApp.getDefaultCalendar();
      var guests=[],date_evt = form.date.split("/");
      // Ajouts de la liste des invités
      guests.push(form.selectrp);
      guests.push(row.emetteur);
      
      
      
      
      if(__DEBUG__) { Logger.log(date_evt);Logger.log('Default calendar is "%s" .guests():%s  .date:%s', calendar.getName(),guests.join(),new Date(date_evt[2],date_evt[1]-1,date_evt[0]));}
      calendar.createAllDayEvent('Go/No go: '+row.client+'-'+row.dossier+' à confirmer',
                                 new Date(date_evt[2],date_evt[1]-1,date_evt[0]),
        { location: 'SQLI Agence St denis',
          description:'merci de confirmer le rdv avec la bonne liste de diffusion des participants ainsi que la salle',
          guests:guests.join(),
          sendInvites:true});
      
      // Envoie du mail au Sales
      MailingToSales(row.emetteur,{client:row.client,
                                   ao:row.dossier,
                                   //Begin Bug Fix MANTIS 13
                                   unit:row_b.unit,
                                   // End Bug Fix
                                   date:(row_b.remise===undefined )?'':Utilities.formatDate(row_b.remise, "GMT+02:00", "dd/MM/yyyy"),
                                   diffusion:(row.dateheureMission===undefined )?'':Utilities.formatDate(row.dateheureMission, "GMT+02:00", "dd/MM/yyyy"),
                                   bodymail:form.commentaire_sale,
                                   rp:form.selectrp,
                                   go:form.date,
                                   url:row.url,
                                   id:row.identifiant
                                  });
      
      // Envoie du mail au RP
      MailingToRP(form.selectrp,{client:row.client,
                                 ao:row.dossier,
                                 unit:row_b.unit,
                                 date:(row_b.remise===undefined )?'':Utilities.formatDate(row_b.remise, "GMT+02:00", "dd/MM/yyyy"),
                                 diffusion:(row.dateheureMission===undefined )?'':Utilities.formatDate(row.dateheureMission, "GMT+02:00", "dd/MM/yyyy"),
                                 go:form.date,
                                 url:row.url,
                                 id:row.identifiant,
                                 sales:row.emetteur,
                                 /*
                                 BugFix : MANTIS 0032
                                 BUGFIX : Erreur sur call MailingToRP avec manager:row.managerEnCharge. On gère l'utilisation dans ce cas
                                 du champ managerPressenti
                                 */
                                 manager:manager
                                });
      
      
      ret=true;
    }
  
/*\/\/\/\/\/\/\//\/\/\/\/
Begin Bug Fix MANTIS 15
\/\/\/\/\/\/\//\/\/\/\/*/
    if(ret==null) throw 'le dossier r&eacute;f&eacute;rence <b>'+form.identifiant+'</b> est d&eacute;j&agrave; affect&eacute; au responsable de proposition <b>'+row.rp+'</b><br>A bientôt,';
/*\/\/\/\/\/\//\/\/\/\/\/
End Bug Fix MANTIS 15
\/\/\/\/\/\/\//\/\/\/\/*/
    return ret;
   }catch(e){treatmentException_(e)} 
}


function onValidateDecision(form){
 try{ 
  if(form===undefined ) throw 'null form';
  if(__DEBUG__) {Logger.finest('onValidateDecision :%s',JSON.stringify(form,null,'\t'));}
  
  //form.visa, form.dar, form.remise, form.drive_id_files
  var ret=null,state=GetRowParams(COLUMN_STATE);
  
  var objet = new SHEET_AO();
  var row = objet.Info_A(form.reference);
  var row_b = objet.Info_B(form.reference);
  var objet_log= new OLogger(row.log);
  // row.rowNumber, row_b.rowNumber
  /*\/\/\/\/\/\/\/\/\/\/\/\/\/\/
  SAUVEGARDE DES DIFFERENTES DATES
  \/\/\/\/\/\/\/\/\/\/\/\/\/\/\/*/
  switch(form.result){
    case 'go' :
      objet.sheet.getRange(row.rowNumber,COLUMN_SHEET_STATE).setValue(state.go);
      objet.sheet.getRange(row.rowNumber,COLUMN_SHEET_VISA).setValue(form.visa);
      objet.sheet.getRange(row.rowNumber,COLUMN_SHEET_DAR_PLAN).setValue(form.dar);
      /*
      MANTIS_053
      */
      objet.sheet.getRange(row.rowNumber,COLUMN_SHEET_REMISE).setValue(form.remise);
      objet.sheet.getRange(row.rowNumber,COLUMN_SHEET_COMMENT).setValue(form.commentaire);
      objet.sheet.getRange(row.rowNumber,COLUMN_SHEET_LOG).setValue(objet_log.pushSession('Décision GO',state.go,{client:form.client,
                                                                                                           dossier:form.dossier,
                                                                                                           responsable: form.rp,
                                                                                                           remise:form.remise,
                                                                                                           visa:form.visa,
                                                                                                           dar:form.dar,
                                                                                                           synthèse:form.commentaire,}));
      // Push Event Calendar
      var calendar = CalendarApp.getDefaultCalendar();
      var guests=[],date_evt = form.dar.split("/");
      // Ajouts de la liste des invités
      
      guests.push(row.managerEnCharge);
      guests.push(row.emetteur);
      guests.push(GetEMAIL_NOTIFICATION());
      
      
      if(__DEBUG__) { Logger.finest(date_evt);Logger.finest('Default calendar is "%s" .guests():%s  .date:%s', calendar.getName(),guests.join(),new Date(date_evt[2],date_evt[1]-1,date_evt[0]));}
      
      calendar.createAllDayEvent('DAR AVV: '+row.client+'-'+row.dossier.substr(0,10)+'...',
                                 new Date(date_evt[2],date_evt[1]-1,date_evt[0]),
        { location: 'SQLI Agence St denis',
          
          /*
          Modification : MANTIS 0031
          */
          
          description:'Dossier de suivi: '+row.dossier+'\r'+
                      'Merci de confirmer le rdv avec la bonne liste de diffusion des participants dès que la date définitive et l\'horaire seront connus.\r'+
                      'Pour rappel les représentants direction à inviter pour le DAR AVV:\r'+
                      'Budget < 300 k€ : Nicolas Larousse ou Stéphane Escoubes\r'+
                      'Budget compris entre 300 k€ et 1 M€ : Philippe Fuhr\r'+
                      '> 1 M€ : Thierry Chemla ou Didier Fauque, Philippe Fuhr\r'+
                      'il est indispensable d\'organiser un pré-DAR au minimum une semaine avant remise avec la direction de l\'agence. Le DAR final avec la direction SQLI sera à organiser manuellement',
          guests:guests.join(),
          sendInvites:true});
                                                                                                           
                                                                                                           
      break;
    case 'nogo':
      objet.sheet.getRange(row.rowNumber,COLUMN_SHEET_STATE).setValue(state.nogo);
      objet.sheet.getRange(row.rowNumber,COLUMN_SHEET_COMMENT_NOGO).setValue(form.commentaire);
      objet.sheet.getRange(row.rowNumber,COLUMN_SHEET_LOG).setValue(objet_log.pushSession('Décision NOGO',state.nogo,{client:form.client,
                                                                                                           dossier:form.dossier,
                                                                                                           synthèse:form.commentaire,}));
      break;
    default: throw new Error('Traitement sur decision inconnu !');
      
  }
  
  
  
  /*\/\/\/\/\/\/\/\/\/\/\/\/\/\/
  SAUVEGARDE DU FICHIER CR GO-NOGO:
  url googledrive du folder de stockage
  \/\/\/\/\/\/\/\/\/\/\/\/\/\/\/*/
  if( form.drive_id_files) url_drive = saveFiles(form.drive_id_files.split(","),form.reference,ID_CR_FOLDER);
  
  // Préparation & diffusion des mails aux destinataires
  var mail_text;
  
  
  if(form.result=='go') {
    /*
      Lorsqu'un Go ou No Go est prononcé, il faut le diffuser.
      Si Go, mail à
      To : sales du dossier, RP du dossier, manager en charge
      Cc : paris.manager@sqli.com, paris.production
      Texte du mail
      Un Go a été prononcé sur ce dossier
      Commentaire : commentaire saisi par le RP
      avec rappel du nom du RP, date remise, date visa forfait, date DAR
   */
    mail_text='<p>Un Go a été prononcé sur ce dossier</p>'
    
    mail_text+='<p>date de remise:&#09;<b>'+form.remise+'</b></p>'
    mail_text+='<p>date du visa forfait:&#09;<b>'+form.visa+'</b></p>'
    mail_text+='<p>date du DAR:&#09;<b>'+form.dar+'</b></p>'
    mail_text+='<p>RP:&#09;<b>'+row.rp+'</b></p>'
    mail_text+='<hr>'
    mail_text+='<p>Commentaire:</p>'
    
    mail_text+='<p>'+form.commentaire+'</p>'
    mailingAfterGoNoGo(row.rp+','+row.emetteur+','+row.managerEnCharge,
                       { result:'GO',
                         ao:row.dossier,
                         client:row.client,
                         remise:form.remise,
                         rp:row.rp,
                         title:'',
                         comment:mail_text,
                         cc:GetEMAIL_DIFFUSION_GO_CC(),
                       });
    
    
    
    /*
       Envoie par mail d'une demande de visa forfait et de création d'arbo AVV
    */
    MailingToAdminAskVisa(GetEMAIL_ADMIN(),{
                                 id:form.reference,
                                 visa:form.visa,
                                 ao:form.dossier,
                                 client:form.client,
                                 remise:form.remise,
                                 rp:form.rp,
   
                                });
  }else{  // S'agit d'un no go
    /*
      Si No Go, mail à
      To : sales du dossier, RP du dossier, manager en charge, vdupeux
      Cc : paris.manager@sqli.com, paris.production
      Texte du mail
      Un No Go a été prononcé sur ce dossier
      Commentaire : commentaire saisi par le RP
     */
      mail_text='<p>Un No Go a été prononcé sur ce dossier</p>'
      
      mail_text+='<p>Commentaire:</p>'
      
      mail_text+='<p>'+form.commentaire+'</p>'
     
      mailingAfterGoNoGo(row.rp+','+row.emetteur+','+row.managerEnCharge+','+GetEMAIL_DIFFUSION_NOGO_TO(),
                       { result:'NO GO',
                         ao:row.dossier,
                         client:row.client,
                         remise:form.remise,
                         rp:row.rp,
                         title:'',
                         comment:mail_text,
                         cc:GetEMAIL_DIFFUSION_NOGO_CC(),
                       });
  
  
  }
  
  return true;
  }catch(e){treatmentException_(e)} 
}


function GetRowParams(header){
 try{
  var sheet = SpreadsheetApp.openById(CLSID).getSheetByName(SHEET_PARAMS);
  
  
  var column ={pos:-1};
  var values = sheet.getDataRange()
                    .getValues();
  
  for( var i=0;i< values[0].length;i++){
    if(values[0][i]==header ) {column.pos=i+1;} 
  }
  
  for( var i in column){ if(column[i]==-1) throw 'no state param';}
  
  var params = sheet.getRange(2, column.pos, MAX_ROW, 1)
                    .getValues();
                    
  params=  params.transpose();

  
  return getObjects(params, normalizeHeaders(params[0]) )[0];
  
  }catch(e){treatmentException_(e)}   
  
}

function enumParams(header){ 
try{
var pos = SpreadsheetApp.openById(CLSID).getSheetByName(SHEET_PARAMS)
                                      .getRange(1, 1, 1, SpreadsheetApp.openById(CLSID).getSheetByName(SHEET_PARAMS).getLastColumn() )
                                      .getValues()[0]
                                      .lastIndexOf(header)+1;


return cleanArray(SpreadsheetApp.openById(CLSID).getSheetByName(SHEET_PARAMS)
                                         .getRange(2, pos, MAX_ROW, 1 )
                                         .getValues()
                                         .transpose()[0]);
                                         
}catch(e){treatmentException_(e)}   
}



/**
 * Gets the user's OAuth 2.0 access token so that it can be passed to Picker.
 * This technique keeps Picker from needing to show its own authorization
 * dialog, but is only possible if the OAuth scope that Picker needs is
 * available in Apps Script. In this case, the function includes an unused call
 * to a DriveApp method to ensure that Apps Script requests access to all files
 * in the user's Drive.
 *
 * @return {string} The user's OAuth 2.0 access token.
 */
function getOAuthToken() {
  try{
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
  
  }catch(e){treatmentException_(e)}   
}



function buildHtmlTable_48h(){
try{ 
  var html ='';
  
  var ret=null, sheet =  SpreadsheetApp.openById(CLSID)
                           .getSheetByName(SHEET_O_AVV),
      sheet_form = SpreadsheetApp.openById(CLSID)
                           .getSheetByName(SHEET_O_HISTO);
      
  var lastRow = sheet.getLastRow(), state=GetRowParams(COLUMN_STATE);
  
  var data =getRowsData(sheet, sheet.getRange(2, 1, sheet.getMaxRows() - 1, COLUMN_SHEET_MESSAGE));
  var values_ao =sheet_form.getDataRange().getValues();
  var data_ao = getObjects(values_ao, normalizeHeaders(values_ao[0]));
  var values_ao_search = values_ao.transpose()[0];
   
  html+='<table class="layout display responsive-table">';
  html+='<thead><tr>';
  html+= '<th>Client-Dossier</th>';
  html+= '<th>Sales</th>';
  html+= '<th>CRM</th>';
  html+= '<th>Emis le</th>';
  
  html+='<th>Retard<br><small>ouvré</small></th>';
  html+='<th>Manager<br><small> pressenti</small></th>';
  html+='<th>Manager<br><small> redirigé</small></th>';
  html+='<th>Relance<br><small> dernière date</small></th>';
  html+='<th>Actions</th>';
  html+='</thead></tr>';
  html+='<tbody>';
  
  for (var i = 0; i < data.length; ++i) {
    var row = data[i],manager='';
    row.rowNumber = i+2;
    if((row.statut==state.enAttente || row.statut==state.relance) && (row.nbJoursRetard > MAX_DAY_NO_ENCOURS_FOR_ALERTE )){
      /*\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\//\/\/\/\/\/
      Recherche des infos liée à la fiche AO
      \/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\//\/\/\/\/\/*/
      var result_search = values_ao_search.indexOf(row.identifiant);
      var row_ao = (result_search!=-1)?data_ao[result_search]:null;
      
      html+='<tr>';
      
      html+='<td class="organisationnumber" data-title="">'+row.client+'-'+row.dossier+'</td>';
      html+='<td class="organisationname" data-title="Sales:">'+row.emetteur.split('@')[0]+'</td>';
      
      html+='<td class="organisationname" data-title="CRM:">'+(row_ao?row_ao.crm:CORRUPT_DATA)+'</td>';
      html+='<td class="organisationname" data-title="Emis le:">'+getDate(row.dateheureMission)+'</td>';
      
      html+=(row.nbJoursRetard > MAX_DAY_NO_ENCOURS_FOR_ALERTE)?'<td class="retard warning" data-title="retard:">'+row.nbJoursRetard+'</td>':'<td class="retard" data-title="retard:">'+row.nbJoursRetard+'</td>';
      
      if(row.managerPressenti!==undefined && row.managerPressenti.indexOf('@')!=-1){
        html+='<td class="organisationname" data-title="Pressenti:">'+row.managerPressenti.split('@')[0]+'</td>';
      }
      else
        html+='<td class="organisationname" data-title="Pressenti:">&nbsp;</td>';
      
      if(row.managerEnCharge!==undefined && row.managerEnCharge.indexOf('@')!=-1){
        html+='<td class="organisationname" data-title="Redirigé:">'+row.managerEnCharge.split('@')[0]+'</td>';
      }
      else
        html+='<td class="organisationname" data-title="Redirigé:">&nbsp;</td>';
      
      html+='<td class="organisationname" data-title="Relance:">'+'dd/mm/aaaa'+'</td>';
      html+='<td class="actions" ><a href="#" class="edit-item" title="Edit" id="relance"><span class="glyphicon glyphicon-share" aria-hidden="true" style="color: #739931;" > Relance</span></a></td>';
      
      html+='</tr>';
    }
  }
  
    
    
  html+='</tbody>';
  html+='</table>';
  
  if(__DEBUG__) {Logger.finest(html);}
  
  return html;
  
}catch(e){treatmentException_(e)}   
    
}

