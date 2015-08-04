

/*

http://templates.mailchimp.com/resources/inline-css/

*/

function testingMail(){
  try{
    
        
    mailingAfterGoNoGo('eremy@sqli.com',{ 
                         ao:'dossier',
                         client:'Mon client',
                         remise:"10/10/2015",
                         rp:'toto@sqli.com',
                         title:'Mon titre ici',
                         comment:'<p>VOILA</p>',
                         cc:GetEMAIL_DIFFUSION_NOGO_CC(),
                       });
    
  }catch(e){Logger.log(e);}

}

function mailingAfterGoNoGo(destinataire,option){
  if(__DEBUG__) {Logger.log('MailingToAdminAskVisa to:%-20s\toption:%s',destinataire,JSON.stringify(option,null,null));}
  if(destinataire===undefined || typeof option!='object') throw 'MailingToAdminAskVisa:no values for sendMailing';
  
  var html = HtmlService.createTemplateFromFile('mailchimp_standard_cssonline'),
      objet='',mail='';

  html.client= option.client;
  html.ao = option.ao;
  html.body_html= '%%body_mail%%';
  
  
  var sqliLogoUrl = "http://www.sqli-enterprise.com/files/2014/05/logo_sqli_entreprise_340x156_bg_transp1.png";
   
  var sqliLogoBlob = UrlFetchApp
                          .fetch(sqliLogoUrl)
                          .getBlob()
                          .setName("sqliLogoBlob");
  
  var advancedArgs = {htmlBody: html.evaluate()
                                    .getContent()
                                    .replace('%%body_mail%%',option.comment.nl2br()),
                                    
                      inlineImages:{
                                     sqliLogo: sqliLogoBlob
                                   },
                      cc: option.cc,
                     
                     };
  
  objet = SUBJECT_DECISION_GONOGO.replace('%%client%%',html.client)
  .replace('%%result%%',option.result)
  .replace('%%ao%%',html.ao)
  .replace('%%date%%',option.remise)
  .replace('%%unit%%','(RP:'+option.rp+')');
  
  
  if(option) MailApp.sendEmail(destinataire, objet, '', advancedArgs);
  
  
  
  
  return true;

}


function MailingToAdminAskVisa(destinataire,option){
  if(__DEBUG__) {Logger.log('MailingToAdminAskVisa to:%-20s\toption:%s',destinataire,JSON.stringify(option,null,null));}
  if(destinataire===undefined || typeof option!='object') throw 'MailingToAdminAskVisa:no values for sendMailing';
  
  var html = HtmlService.createTemplateFromFile('mailchimp_standard_cssonline'),
      objet='',mail='';
  
  html.client= option.client;
  html.ao = option.ao;
  html.body_html= '%%body_mail%%';
  
  mail = '<p>Une demande de <b>Visa Forfait</b> pour le:<br></p>';
  mail+='<p align="center" style="font-size:1.2em;"><b>'+option.visa+'</b></p>';
  mail+='<p>Une demande de création d\'arborescence.</p>';
  
  mail+= '<table border="0" cellpadding="0" cellspacing="0" style="text-align:center;margin-left:10%;float:left;background-color:#0099CC; border:1px solid #006699; border-radius:5px;">'
  mail+= '<tr>'
  mail+= '<td align="center" valign="middle" style="color:#FFFFFF; font-family:Helvetica, Arial, sans-serif; font-size:16px; font-weight:normal; letter-spacing:-.5px; line-height:150%; padding-top:15px; padding-right:30px; padding-bottom:15px; padding-left:30px;">'
  mail+= '<a href="'+URL_EXEC+'?page=build_avv&ref='+option.id+'" target="_blank" style="color:#FFFFFF; text-decoration:none;">Créer l\'arbo GDrive</a>'
  mail+= '</td>'
  
  mail+= '</tr>'
  mail+= '</table>'
  mail+= '<br style="clear:both;"><p ><small>Contacter le support si vous rencontrez des problèmes.</small></p>'
  
  var sqliLogoUrl = "http://www.sqli-enterprise.com/files/2014/05/logo_sqli_entreprise_340x156_bg_transp1.png";
   
  var sqliLogoBlob = UrlFetchApp
                          .fetch(sqliLogoUrl)
                          .getBlob()
                          .setName("sqliLogoBlob");
  
  var advancedArgs = {htmlBody: html.evaluate()
                                    .getContent()
                                    .replace('%%body_mail%%',mail.nl2br()),
                      inlineImages:{
                                     sqliLogo: sqliLogoBlob
                                   },
                      /*
                       MANTIS 049
                      */
                      noReply:true,
                     
                     };
  
  objet = SUBJECT.replace('%%client%%',html.client)
  .replace('%%ao%%',html.ao)
  .replace('%%date%%',option.remise)
  .replace('%%unit%%','(RP:'+option.rp+')');
  
  
  if(option) MailApp.sendEmail(destinataire, objet, '', advancedArgs);
  
  
  
  
  return true;
}

function MailingToRP(emetteur,option){
  
  if(__DEBUG__) {Logger.log('MailingToRP to:%-20s\toption:%s',emetteur,JSON.stringify(option,null,null));}
  if(emetteur===undefined || typeof option!='object') throw 'MailingToRP:no values for sendMailing';
  var html = HtmlService.createTemplateFromFile('mailchimp_standard_cssonline'),
      objet='',mail='';
  
  
  
  html.client= option.client;
  html.ao = option.ao;
  html.body_html= '%%body_mail%%';
  
  mail = '<p>Vous êtes positionné par votre manager en tant que Responsable de Proposition sur le dossier suivant :</p>';
  mail+='<p><table><tr style="-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;page-break-inside: avoid;background: #f8e7e7;"><td style="-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;padding: 0;border-collapse: collapse;mso-pattern: black none;padding-top: 5px;padding-bottom: 5px;padding-right: 2px;padding-left: 7px;mso-ignore: padding;color: black;font-weight: bold;font-style: normal;text-decoration: none;font-family: arial,sans-serif;text-align: left;vertical-align: middle;white-space: normal;">Client:</td><td style="padding-left:10px;">'+option.client+'</td></tr>';
  mail+='<tr style="-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;page-break-inside: avoid;background: #f1cccc;"><td style="-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;padding: 0;border-collapse: collapse;mso-pattern: black none;padding-top: 5px;padding-bottom: 5px;padding-right: 2px;padding-left: 7px;mso-ignore: padding;color: black;font-weight: bold;font-style: normal;text-decoration: none;font-family: arial,sans-serif;text-align: left;vertical-align: middle;white-space: normal;">Dossier:</td><td style="padding-left:10px;">'+option.ao+'</td></tr>';
  mail+='<tr style="-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;page-break-inside: avoid;background: #f8e7e7;"><td style="-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;padding: 0;border-collapse: collapse;mso-pattern: black none;padding-top: 5px;padding-bottom: 5px;padding-right: 2px;padding-left: 7px;mso-ignore: padding;color: black;font-weight: bold;font-style: normal;text-decoration: none;font-family: arial,sans-serif;text-align: left;vertical-align: middle;white-space: normal;">Sales:</td><td style="padding-left:10px;">'+option.sales+'</td></tr>';
  mail+='<tr style="-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;page-break-inside: avoid;background: #f1cccc;"><td style="-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;padding: 0;border-collapse: collapse;mso-pattern: black none;padding-top: 5px;padding-bottom: 5px;padding-right: 2px;padding-left: 7px;mso-ignore: padding;color: black;font-weight: bold;font-style: normal;text-decoration: none;font-family: arial,sans-serif;text-align: left;vertical-align: middle;white-space: normal;">Date de diffusion du dossier:</td><td style="padding-left:10px;">'+option.diffusion+'</td></tr>';
  
  mail+='<tr style="-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;page-break-inside: avoid;background: #f8e7e7;"><td style="-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;padding: 0;border-collapse: collapse;mso-pattern: black none;padding-top: 5px;padding-bottom: 5px;padding-right: 2px;padding-left: 7px;mso-ignore: padding;color: black;font-weight: bold;font-style: normal;text-decoration: none;font-family: arial,sans-serif;text-align: left;vertical-align: middle;white-space: normal;">Date de Go/NoGo:</td><td style="padding-left:10px;">'+option.go+'</td></tr>';
  mail+='<tr style="-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;page-break-inside: avoid;background: #f1cccc;"><td style="-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;padding: 0;border-collapse: collapse;mso-pattern: black none;padding-top: 5px;padding-bottom: 5px;padding-right: 2px;padding-left: 7px;mso-ignore: padding;color: black;font-weight: bold;font-style: normal;text-decoration: none;font-family: arial,sans-serif;text-align: left;vertical-align: middle;white-space: normal;">Date de remise:</td><td style="padding-left:10px;">'+option.date+'</td></tr>';
  if(option.url) 
    mail+='<tr style="-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;page-break-inside: avoid;background: #f8e7e7;"><td style="-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;padding: 0;border-collapse: collapse;mso-pattern: black none;padding-top: 5px;padding-bottom: 5px;padding-right: 2px;padding-left: 7px;mso-ignore: padding;color: black;font-weight: bold;font-style: normal;text-decoration: none;font-family: arial,sans-serif;text-align: left;vertical-align: middle;white-space: normal;">Accès aux documents de l\'AO:</td><td style="padding-left:10px;"><a href="'+option.url+'">'+option.id+'</a></td></tr>';
  else 
     mail+='<tr style="-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;page-break-inside: avoid;background: #f1cccc;"><td style="-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;padding: 0;border-collapse: collapse;mso-pattern: black none;padding-top: 5px;padding-bottom: 5px;padding-right: 2px;padding-left: 7px;mso-ignore: padding;color: black;font-weight: bold;font-style: normal;text-decoration: none;font-family: arial,sans-serif;text-align: left;vertical-align: middle;white-space: normal;">Accès aux documents de l\'AO:</td><td style="padding-left:10px;">Aucun document associé</td></tr>';
  mail+='</table></p>';
  mail+= '<p><b><i>Une fois le Go/no Go réalisé</i></b>, vous devez demander la création de l\'arborescence Avant-vente GDrive, ou bien donner les raisons du No Go.</p>'  
  
  mail+= '<table border="0" cellpadding="0" cellspacing="0" style="margin-left:10%;float:left;background-color:#0099CC; border:1px solid #006699; border-radius:5px;">'
  mail+= '<tr>'
  mail+= '<td align="center" valign="middle" style="color:#FFFFFF; font-family:Helvetica, Arial, sans-serif; font-size:16px; font-weight:normal; letter-spacing:-.5px; line-height:150%; padding-top:15px; padding-right:30px; padding-bottom:15px; padding-left:30px;">'
  mail+= '<a href="'+URL_EXEC+'?page=decision&result=go&ref='+option.id+'" target="_blank" style="color:#FFFFFF; text-decoration:none;">Go !<br>Créer l\'arbo GDrive</a>'
  mail+= '</td>'
  
  mail+= '</tr>'
  mail+= '</table>'
  mail+= '<table border="0" cellpadding="0" cellspacing="0" style="margin-right:10%;float:right;background-color:#CC0000; border:1px solid #990000; border-radius:5px;">'
  mail+= '<tr>'
  mail+= '<td align="center" valign="middle" style="color:#FFFFFF; font-family:Helvetica, Arial, sans-serif; font-size:16px; font-weight:normal; letter-spacing:-.5px; line-height:150%; padding-top:15px; padding-right:30px; padding-bottom:15px; padding-left:30px;">'
  mail+= '<a href="'+URL_EXEC+'?page=decision&result=nogo&ref='+option.id+'" target="_blank" style="color:#FFFFFF; text-decoration:none;">No Go<br>Clore l\'avant-vente</a>'
  mail+= '</td>'
		
  mail+= '</tr>'
  mail+= '</table>'
  mail+= '<br style="clear:both;"><p ><small>Ne pas effacer ce mail. Passer par paris.production si vous rencontrez des problèmes.</small></p>'

  var sqliLogoUrl = "http://www.sqli-enterprise.com/files/2014/05/logo_sqli_entreprise_340x156_bg_transp1.png";
   
  var sqliLogoBlob = UrlFetchApp
                          .fetch(sqliLogoUrl)
                          .getBlob()
                          .setName("sqliLogoBlob");
  
  var advancedArgs = {htmlBody: html.evaluate()
                                    .getContent()
                                    .replace('%%body_mail%%',mail.nl2br()),
                      inlineImages:{
                                     sqliLogo: sqliLogoBlob
                                   },
                      /*
                      BugFix : MANTIS 0032 + MANTIS 0036
                      
                      */
                      cc: option.manager+','+GetEMAIL_NOTIFICATION()
                     };
  
  objet = SUBJECT.replace('%%client%%',html.client)
  .replace('%%ao%%',html.ao)
  .replace('%%date%%',option.date)
  .replace('%%unit%%',option.unit);
  
  
  if(option) MailApp.sendEmail(emetteur, objet, '', advancedArgs);
  
  return true;
}

function MailingToSales(emetteur,option){
  
  if(emetteur===undefined || typeof option!='object') throw 'MailingToSales:no values for sendMailing';
  var html = HtmlService.createTemplateFromFile('mailchimp_standard_cssonline'),
      objet='',mail='';
  
  if(__DEBUG__) {Logger.log('MailingToSales to:%-20s\toption:%s',emetteur,JSON.stringify(option,null,null));}
  
  html.client= option.client;
  html.ao = option.ao;
  html.body_html= '%%body_mail%%';
  
  mail = option.bodymail;
  mail+='<br><hr><table><tr><td>Responsable de proposition affecté:</td><td style="padding-left:10px;"><u>'+option.rp+'</u></td></tr>';
  mail+='<tr><td>Date de Go/NoGo:</td><td style="padding-left:10px;"><u>'+option.go+'</u></td></tr>';
  if(option.url) 
    mail+='<tr><td>Accès aux documents de l\'AO sous:</td><td style="padding-left:10px;"><a href="'+option.url+'">'+option.id+'</a></td></tr>';
  else 
     mail+='<tr><td>Aucun document associé</td></tr>';
  mail+='</table>';
 
  var sqliLogoUrl = "http://www.sqli-enterprise.com/files/2014/05/logo_sqli_entreprise_340x156_bg_transp1.png";
   
  var sqliLogoBlob = UrlFetchApp
                          .fetch(sqliLogoUrl)
                          .getBlob()
                          .setName("sqliLogoBlob");
  
  var advancedArgs = {htmlBody: html.evaluate()
                                    .getContent()
                                    .replace('%%body_mail%%',mail.nl2br()),
                      inlineImages:{
                                     sqliLogo: sqliLogoBlob
                                   },
                      /*
                      Modification : MANTIS 0039
                      */
                      
                      cc: GetEMAIL_ADMIN()+','+option.rp
                     };
  
  objet = SUBJECT.replace('%%client%%',html.client)
  .replace('%%ao%%',html.ao)
  .replace('%%date%%',option.date)
  .replace('%%unit%%',option.unit);
  
  
  if(option) MailApp.sendEmail(emetteur, objet, '', advancedArgs);
  
  return true;
}

function MailingToOtherManager(emetteur,option){
  if(emetteur===undefined || typeof option!='object') throw 'MailingToOtherManager:no values for sendMailing';
  var html = HtmlService.createTemplateFromFile('mailchimp_cssonline'),
      objet='',mail='',arbitrage_enum = enumParams(COLUMN_ARBITRAGE);
  
  if(__DEBUG__){ Logger.log('MailingToOtherManager to:%-20s\toption:%s',emetteur,JSON.stringify(option,null,null));}

  html.client= option.client;
  html.ao = option.ao;
  html.contexte = '%%contexte%%' //option.contexte;
  html.enjeux = '%%enjeux%%' //option.enjeux;
  html.unit = option.unit;
  html.budget = option.budget;
  html.date = getDate(option.date);
  html.ic = option.indice;
  html.criteres = '%%critere%%'   //option.critere;
  html.origine = '%%origine%%'   //option.origine;
  html.pourquoi = '%%pourquoi%%'  //option.pourquoi;
  html.partenaires = option.partenaires;
  html.concurrents = option.concurrents;
  html.crm = option.crm;
  html.sales = option.emetteur;
  html.techno = option.techno;
  
  html.form_response = (arbitrage_enum.indexOf(html.unit)!=-1)?URL_EXEC+'?page=response&type=expanded&ref='+option.identifiant:URL_EXEC+'?page=response&ref='+option.identifiant;
  
  //html.form_response = URL_EXEC+'?page=response&ref='+option.identifiant;
  html.form_response = encodeURI(html.form_response);
  if(__DEBUG__){ Logger.log('MailingToOtherManager:%s',html.form_response);}
  
  html.manager = option.manager;
  
  html.body_html= '%%body_mail%%';
  
  mail = (function ALERTE_MESSAGE_INLINE(msg) {return '<div style="text-decoration:none;-webkit-box-sizing: border-box;-moz-box-sizing: border-box;box-sizing: border-box;padding: 15px;margin-bottom: 20px;border: 1px solid transparent;border-radius: 4px;color: #a94442;background-color: #f2dede;border-color: #ebccd1;margin-right: 15px;margin-left: 15px;"><b>&#9432;</b>'+msg+'</div>';})
                                            (' Le manager pressenti pour répondre à ce dossier (<b>'+option.previous_manager+'</b>), l\'a redirigé vers <b>'+option.manager+'</b>'); 
  mail += option.bodymail;
  mail+='Message complementaire de <a href="mailto:'+option.previous_manager+'">'+option.previous_manager+':</a><br>'+option.info_complementaire+'<hr>';
    
  var sqliLogoUrl = "http://www.sqli-enterprise.com/files/2014/05/logo_sqli_entreprise_340x156_bg_transp1.png";
   
  var sqliLogoBlob = UrlFetchApp
                          .fetch(sqliLogoUrl)
                          .getBlob()
                          .setName("sqliLogoBlob");
  
  var advancedArgs = {htmlBody: html.evaluate()
                      .getContent()
                      .replace('%%body_mail%%',(mail==undefined)?'':mail.nl2br())
                      .replace('%%contexte%%',(option.contexte==undefined)?'':option.contexte.nl2br())
                      .replace('%%enjeux%%',(option.enjeux==undefined)?'':option.enjeux.nl2br())
                      .replace('%%critere%%',(option.critere==undefined)?'':option.critere.nl2br())
                      .replace('%%origine%%',(option.origine==undefined)?'':option.origine.nl2br())
                      .replace('%%pourquoi%%',(option.pourquoi==undefined)?'':option.pourquoi.nl2br()),
                      inlineImages:{
                                     sqliLogo: sqliLogoBlob
                                   },
                      cc: EMAIL_CC +','+emetteur
                     };
  objet = SUBJECT.replace('%%client%%',html.client)
  .replace('%%ao%%',html.ao)
  .replace('%%date%%', html.date)
  .replace('%%unit%%',html.unit);

  if(option) MailApp.sendEmail(option.manager, objet, '', advancedArgs);

  return true;
}

function MailingToManager(emetteur,id,values,mail,option){
  
  if(values===undefined || typeof values!='object') throw 'MailingToManager:no values for sendMailing';
  var html = HtmlService.createTemplateFromFile('mailchimp_cssonline'),objet='',
      arbitrage_enum = enumParams(COLUMN_ARBITRAGE);
  
  
  Logger.log('MailingToManager to:%-20s\toption:%s',emetteur,JSON.stringify(values,null,null));
  
  html.client= values.find('client');
  html.ao = values.find('ao');
  html.contexte = '%%contexte%%'   //values.find('contexte').nl2br();
  html.enjeux = '%%enjeux%%'       //values.find('enjeux').nl2br();
  html.unit = values.find('unit');
  html.budget = values.find('budget');
  html.date = values.find('date');
  html.ic = values.find('indice');
  html.criteres = '%%critere%%'    //values.find('critere');
  html.origine = '%%origine%%'     // values.find('origine');
  html.pourquoi = '%%pourquoi%%'   //values.find('pourquoi');
  html.partenaires = values.find('partenaires');
  html.concurrents = values.find('concurrents');
  html.crm = values.find('crm').toUpperCase();
  html.sales = emetteur;
  html.url_drive = 'url drive';
  html.techno= values.find('techno')+'-'+values.find('techno-autre');
  html.form_response = (arbitrage_enum.indexOf(html.unit)!=-1)?URL_EXEC+'?page=response&type=expanded&ref='+id:URL_EXEC+'?page=response&ref='+id;
  
  html.form_response = encodeURI(html.form_response);
  html.manager = GetManager(values.find('unit'));
  
  html.body_html='%%body_mail%%';
  
  
  var sqliLogoUrl = "http://www.sqli-enterprise.com/files/2014/05/logo_sqli_entreprise_340x156_bg_transp1.png";
   
  var sqliLogoBlob = UrlFetchApp
                          .fetch(sqliLogoUrl)
                          .getBlob()
                          .setName("sqliLogoBlob");
  
  var advancedArgs = {htmlBody: html.evaluate()
                      .getContent()
                      .replace('%%body_mail%%',(mail==undefined)?'':mail.nl2br())
                      .replace('%%contexte%%',values.find('contexte').nl2br())
                      .replace('%%enjeux%%',values.find('enjeux').nl2br())
                      .replace('%%critere%%',values.find('critere').nl2br())
                      .replace('%%origine%%',values.find('origine').nl2br())
                      .replace('%%pourquoi%%',values.find('pourquoi').nl2br()),
                      inlineImages:{
                                     sqliLogo: sqliLogoBlob
                                   },
                      cc: EMAIL_CC +','+emetteur
                     };
  objet = SUBJECT.replace('%%client%%',html.client)
  .replace('%%ao%%',html.ao)
  .replace('%%date%%',html.date)
  .replace('%%unit%%',html.unit);
  
  if(option) MailApp.sendEmail(html.manager, objet, '', advancedArgs);
  
  return advancedArgs;
}

function MailContact(str){
   MailApp.sendEmail( GetEMAIL_CONTACT(), 'PP9-WEB: Questions utilisateurs', str );
}


String.prototype.nl2br = function()
{
    return this.replace(/(\n)/g, "<br />");
}

