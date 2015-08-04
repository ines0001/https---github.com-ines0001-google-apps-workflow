function testing_notif(){
  
 Logger.log(enumParams(COLUMN_ARBITRAGE).indexOf('Moblité') )
  
}

/**
 * Construction du mail de notifications sur 
 * dossiers sans RP avec UNIT sur demande d'arbitrage
 * @mail {diffusion [clé Arbitrage UNIT params]}.
 * Déclenchable par Trigger
 */
function MailingToAdminForArbitrage(){
  try{
  var html = HtmlService.createTemplateFromFile('mailchimp_standard_cssonline');
  var SUBJECT_MAIL ='AVV en arbitrage du %date%'
  
  html.client= 'AVV en arbitrage';
  html.ao ='';
  html.body_html = '%%TAG%%';
  
   
  var sqliLogoBlob = UrlFetchApp
                          .fetch("http://www.sqli-enterprise.com/files/2014/05/logo_sqli_entreprise_340x156_bg_transp1.png")
                          .getBlob()
                          .setName("sqliLogoBlob");
  
 
  var advancedArgs = {htmlBody: html.evaluate()
                                    .getContent()
                                    .replace('%%TAG%%',buildHtmlTableNotificationArbitrage()),
                      inlineImages:{
                                     sqliLogo: sqliLogoBlob
                                   },
                      cc: Session.getActiveUser().getEmail()
                     };
  
  
  MailApp.sendEmail(GetEMAIL_ADMIN(), SUBJECT_MAIL.replace('%date%',getDate()),'', advancedArgs);
  }catch(e){Logger.log(e);}
  
  
}

/**
 * Construction du mail de notifications sur 
 *  dossiers sans RP avec UNIT sur demande d'arbitrage
 * @mail {diffusion [clé Arbitrage UNIT params]}.
 * Déclenchable par Trigger
 */

function buildHtmlTableNotificationArbitrage(){
  var html ='',state=GetRowParams(COLUMN_STATE), arbitrage_enum = enumParams(COLUMN_ARBITRAGE);
  
  var ret=null, sheet =  SpreadsheetApp.openById(CLSID)
                           .getSheetByName(SHEET_O_AVV),
      sheet_form = SpreadsheetApp.openById(CLSID)
                           .getSheetByName(SHEET_O_HISTO);
 
  var data =getRowsData(sheet, sheet.getRange(2, 1, sheet.getMaxRows() - 1, COLUMN_SHEET_MESSAGE));

  
  var values_ao =sheet_form.getDataRange().getValues();
  var data_ao = getObjects(values_ao, normalizeHeaders(values_ao[0]));
  var values_ao_search = values_ao.transpose()[0];
  
  
  for (var i = 0; i < data.length; ++i) {
    var row = data[i],manager='';
    /*\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\//\/\/\/\/\/
    Recherche des infos liée à la fiche AO
    \/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\//\/\/\/\/\/*/
    var result_search = values_ao_search.indexOf(row.identifiant);
    var row_ao = (result_search!=-1)?data_ao[result_search]:null;
    
    
    /*\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\//\/\/\/\/\/
    Information à historiser
     
    Historiser=(function ALERTE_MESSAGE_INLINE(msg) {return '<div style="text-decoration:none;-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;padding: 15px;margin-bottom: 20px;border: 1px solid transparent;border-radius: 4px;color: #a94442;background-color: #f2dede;border-color: #ebccd1;margin-right: 15px;margin-left: 15px;"><b>&#9432;</b>'+msg+'</div>';})
                                            ('<b>Dossier :</b>'+row.dossier+' <b>Client:</b>'+row.client+' <b>Réf:</b>'+row.identifiant+' <b>Error: Fiche incomplète !</b>');
    \/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\//\/\/\/\/\/*/
    
    if( row_ao && (row.statut==state.enAttente || row.statut==state.relance)&& arbitrage_enum.indexOf(row_ao.unit)!=-1 ){
      
      
      html+='<table class="layout display" style="width: 100%;border-collapse: collapse;margin: 1em 0;box-shadow: 0 1px 10px rgba(0, 0, 0, 0.2);">';
    html+='<thead>';
        html+='<tr>';
            html+='<th width="25%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #f2dede;"><span class="retard" style="min-width: 15px;padding: 5px 5px;font-size: 13px;font-weight: bold;line-height: 1;text-align: center;white-space: nowrap;vertical-align: baseline;border-radius: 10px;color: #fff;background-color: #333;">Arbitrage</span></th>';
            html+='<th width="65%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #f2dede;">'+row.client+'-'+row.dossier+'</th>';
            html+='<th width="10%" style="text-align: right;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #f2dede;"><a href="'+URL_EXEC+'?page=response&type=expanded&ref='+row.identifiant+'" style="color: #a94442;">affecter</a></th>';
         html+='</tr>';
     html+='</thead>';
  html+='<tbody>';  
     html+='<tr>';
      html+='<td width="25%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">Sales:</td>';
      html+='<td width="65%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">'+row.emetteur.split('@')[0]+'</td>';
      html+='<td width="10%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;"></td>';
     html+='</tr>';
     html+='<tr>';
      html+='<td width="25%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">CRM:</td>';
      html+='<td width="65%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">'+(row_ao?row_ao.crm:CORRUPT_DATA)+'</td>';
      html+='<td width="10%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;"></td>';
     html+='</tr>';
     html+='<tr>';
      html+='<td width="25%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">Emis le:</td>';
      html+='<td width="65%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">'+getDate(row.dateheureMission)+'</td>';
      html+='<td width="10%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;"></td>';
     html+='</tr>';
     html+='<tr>';
      html+='<td width="25%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">Manager <small>pressenti</small></td>';
      
      html+='<td width="65%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">'+row_ao.unit+'</td>';
      
     
     html+='</tr>';
     
    
     
    html+='</tbody>';
 html+='</table>';
      
    } // End If
    
  } // End boucle For
  
  /*\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\//\/\/\/\/\//\/\/\/\/\//\/\/\/\/\//\/\/\/\/\//\/\/\/\/\/
  Vérifier si aucun dossier afficher. Si pas de dossier alors un message de synthèse
  \/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\//\/\/\/\/\//\/\/\/\/\//\/\/\/\/\//\/\/\/\/\//\/\/\/\/\/*/
  if(html=='') html+=(function ALERTE_MESSAGE_INLINE(msg) {return '<div style="text-decoration:none;-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;padding: 15px;margin-bottom: 20px;border: 1px solid transparent;border-radius: 4px;color: #a94442;background-color: #f2dede;border-color: #ebccd1;margin-right: 15px;margin-left: 15px;"><b>&#9432;</b>'+msg+'</div>';})
                                            (' Aucun arbitrage en cours !');
      
  
  return html;
}

/**
 * Construction du mail de notifications sur 
 * dossiers sans RP avec délais > 48h suite à la demande du sales
 * @mail {diffusion [clé Admin params]}.
 * Déclenchable par Trigger
 */

function MailingToAdmin(){
  try{
  var html = HtmlService.createTemplateFromFile('mailchimp_standard_cssonline');
  
  html.client= 'Dossiers non traités';
  html.ao ='délais >48h';
  html.body_html = '%%TAG%%';
  
   
  var sqliLogoBlob = UrlFetchApp
                          .fetch("http://www.sqli-enterprise.com/files/2014/05/logo_sqli_entreprise_340x156_bg_transp1.png")
                          .getBlob()
                          .setName("sqliLogoBlob");
  
 
  var advancedArgs = {htmlBody: html.evaluate()
                                    .getContent()
                                    .replace('%%TAG%%',buildHtmlTableNotification()),
                      inlineImages:{
                                     sqliLogo: sqliLogoBlob
                                   },
                      cc: Session.getActiveUser().getEmail()
                     };
  
  
  MailApp.sendEmail(GetEMAIL_ADMIN(), SUBJECT_ADMIN.replace('%date%',getDate()),'', advancedArgs);
  }catch(e){Logger.log(e);}
}


/**
 * Construction de la table des alertes par notifications 
 * liste des dossiers ouverts sans réponse d'affectation de rp
 * délais > 48h suite à la demande du sales
 * @return {html} table responsive.
 */

function buildHtmlTableNotification(){
  var html ='',state=GetRowParams(COLUMN_STATE);
  
  var ret=null, sheet =  SpreadsheetApp.openById(CLSID)
                           .getSheetByName(SHEET_O_AVV),
      sheet_form = SpreadsheetApp.openById(CLSID)
                           .getSheetByName(SHEET_O_HISTO);
 
  var data =getRowsData(sheet, sheet.getRange(2, 1, sheet.getMaxRows() - 1, COLUMN_SHEET_MESSAGE));

  
  var values_ao =sheet_form.getDataRange().getValues();
  var data_ao = getObjects(values_ao, normalizeHeaders(values_ao[0]));
  var values_ao_search = values_ao.transpose()[0];
  
  
  for (var i = 0; i < data.length; ++i) {
    var row = data[i],manager='';
    if( row.nbJoursRetard > MAX_DAY_NO_ENCOURS_FOR_ALERTE && (row.statut==state.enAttente || row.statut==state.relance) ){
      /*\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\//\/\/\/\/\/
      Recherche des infos liée à la fiche AO
      \/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\//\/\/\/\/\/*/
      var result_search = values_ao_search.indexOf(row.identifiant);
      var row_ao = (result_search!=-1)?data_ao[result_search]:null;
          
html+='<table class="layout display" style="width: 100%;border-collapse: collapse;margin: 1em 0;box-shadow: 0 1px 10px rgba(0, 0, 0, 0.2);">';
    html+='<thead>';
        html+='<tr>';
            html+='<th width="25%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #D5E0CC;"><span class="retard" style="min-width: 15px;padding: 5px 5px;font-size: 13px;font-weight: bold;line-height: 1;text-align: center;white-space: nowrap;vertical-align: baseline;border-radius: 10px;color: #fff;background-color: #333;">retard:'+row.nbJoursRetard+'</span></th>';
            html+='<th width="65%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #D5E0CC;">'+row.client+'-'+row.dossier+'</th>';
            html+='<th width="10%" style="text-align: right;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #D5E0CC;"><a href="'+URL_EXEC+'?page=relance&ref='+row.identifiant+'" style="color: #739931;">relance</a></th>';
         html+='</tr>';
     html+='</thead>';
  html+='<tbody>';  
     html+='<tr>';
      html+='<td width="25%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">Sales:</td>';
      html+='<td width="65%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">'+row.emetteur.split('@')[0]+'</td>';
      html+='<td width="10%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;"></td>';
     html+='</tr>';
     html+='<tr>';
      html+='<td width="25%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">CRM:</td>';
      html+='<td width="65%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">'+(row_ao?row_ao.crm:CORRUPT_DATA)+'</td>';
      html+='<td width="10%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;"></td>';
     html+='</tr>';
     html+='<tr>';
      html+='<td width="25%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">Emis le:</td>';
      html+='<td width="65%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">'+getDate(row.dateheureMission)+'</td>';
      html+='<td width="10%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;"></td>';
     html+='</tr>';
     html+='<tr>';
      html+='<td width="25%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">Manager <small>pressenti</small></td>';
      if(row.managerPressenti!==undefined && row.managerPressenti.indexOf('@')!=-1){
        html+='<td width="65%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">'+row.managerPressenti.split('@')[0]+'</td>';
      }
      else
        html+='<td width="65%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">&nbsp;</td>';
      
      html+='<td width="10%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;"></td>';
     html+='</tr>';
     html+='<tr>';
      html+='<td width="25%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">Manager <small>redirigé</small></td>';
      if(row.managerEnCharge!==undefined && row.managerEnCharge.indexOf('@')!=-1){
        html+='<td width="65%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">'+row.managerEnCharge.split('@')[0]+'</td>';
      }
      else
        html+='<td width="65%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">&nbsp;</td>';
      
      html+='<td width="10%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;"></td>';
     html+='</tr>';
     html+='<tr>';
      html+='<td width="25%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">Relance <small>précédente</small></td>';
      html+='<td width="65%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;">xx/xx/xxxx</td>';
      html+='<td width="10%" style="text-align: left;border-bottom: 1px solid #B3BFAA;padding: .5em;background: #fff;"></td>';
     html+='</tr>';
    
     
    html+='</tbody>';
 html+='</table>';
    }
  }
  
  
  return html;
    
}
