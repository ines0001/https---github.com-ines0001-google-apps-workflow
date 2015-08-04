/*/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
FICHIERS DECLARATION :
  -CONSTANTES GLOBALES
  -FONCTIONS GLOBALES
  -CLASSES
         -----------------------------------
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\*/

/*
Necessaire pour le déploiement de la version
28/07/15: positionner à true afin de produire une nouvelle version
Menu Fichier/Gérer les versions depuis le projet google script apps

28/07/15: 10:17
*/
var __PROD__  = false;

var ID_CLSID_PROD = '1h3y9ferdQq8S6Xk48B01sEsEpC4czP4gdncB9l9AoJY';
//var CLSID =     '1B7TBEOMVfg1kiPGrlW3dY5wWpYdmqOLwDi985L7sKo0';
var ID_CLSID_DEV ='1B7TBEOMVfg1kiPGrlW3dY5wWpYdmqOLwDi985L7sKo0';
var CLSID = (__PROD__)?ID_CLSID_PROD:PropertiesService.getScriptProperties().getProperty('ID_CLSID')



var ID_FOLDER_FILES= '0BxTfS7jXR_FYMWQ4XzhUTDJhS3M';
var ID_CR_FOLDER='CR_FOLDER';


var CORRUPT_DATA='Fiche donnée perdue!';

var SHEET_PARAMS = 'Params';
var SHEET_O_AVV = 'Suivi AO';
var SHEET_O_HISTO = 'Histo AO';
var SHEET_O_MASQUE_ROW = 'Masque Sheet';

/*
Limite de lecture pour une colonne de paramtrage ( Ex : 1000 RP possible ! )
*/
var MAX_ROW      = 1000;
var MAX_DAY_NO_ENCOURS_FOR_ALERTE = 2;


/*/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
    CARTOGRAPHIE des index name de la page de paramétrage
         
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\*/
var INIT_PRACTICE_VALUE ='--Practice--';
var INIT_UNIT_VALUE='--Unit--';
var COLUMN_PRACTICE = 'Practice';
var COLUMN_UNIT = 'Unit';
var COLUMN_UNIT_PRESS = 'UNIT pressentie';
var COLUMN_MANAGER = 'Manager';
var COLUMN_STATE ='Statut';
var COLUMN_RP  ='RP';
var COLUMN_CC  ='Diffusion';
var COLUMN_CONTACT  ='Contact';
var COLUMN_ADMIN  ='Admin';
var COLUMN_FORMULES = 'Formules';
var COLUMN_ARBITRAGE = 'Arbitrage UNIT';
var COLUMN_DASHBOARD_SQL = 'Dashboard Requete';
var COLUMN_DASHBOARD_TABLE_SQL = 'Dashboard table_synthese';
var COLUMN_LIST_TECHNOS = 'Liste des technos'

var COLUMN_NOTIFICATION = 'Notification DAR';  //NE PAS OUBLIER DE CHANGER LE NOM DU CHAMP PAR LA SUITE
var COLUMN_NOTIFICATION_GO_CC ='Diffusion go (cc)' ; // Liste de diffusion en cc après un GO envoyé par le RP
var COLUMN_NOTIFICATION_NOGO_CC ='Diffusion nogo (cc)'; // Liste de diffusion en cc après un NOGO envoyé par le RP
var COLUMN_NOTIFICATION_NOGO_TO ='Diffusion nogo (to)'; // Liste de diffusion  après un NOGO envoyé par le RP (ex: vdupeux)


/*/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
    CARTOGRAPHIE des positions de la page AOSuivi & HistoAO
         
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\*/
var COLUMN_SHEET_MANAGER = 7;

var COLUMN_SHEET_DATE_GONOGO = 13;
var COLUMN_SHEET_RP = 12;
var COLUMN_SHEET_RU = 11;
var COLUMN_SHEET_RETARD = 9;
var COLUMN_SHEET_DATE_RESP = 10;
var COLUMN_SHEET_STATE = 15;
var COLUMN_SHEET_COMMENT_NOGO = COLUMN_SHEET_STATE+1;
var COLUMN_SHEET_COMMENT = COLUMN_SHEET_STATE+2;

var COLUMN_SHEET_VISA = COLUMN_SHEET_STATE+3;
var COLUMN_SHEET_DAR_PLAN = COLUMN_SHEET_STATE+7;
var COLUMN_SHEET_REMISE = 8;
var COLUMN_SHEET_UNIT = 4;

var COLUMN_SHEET_MESSAGE = 40;
var COLUMN_SHEET_LOG = 26;
var MAX_COL = 27;
var CEL_FORMULE_RETARD_PARAMS ='M2';

/*/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
    Fonction d'accès à la formule de calcul de retard
       - A externaliser dans un page dédié pour les 
         modèles de lignes   
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\*/
var RANGE_FORMULA_RETARD_PARAMS = (function () {var params = SpreadsheetApp.openById(CLSID).getSheetByName(SHEET_PARAMS).getDataRange().getValues();
                           for( var i=0;i< params[0].length;i++){if(params[0][i]==COLUMN_FORMULES ) return SpreadsheetApp.openById(CLSID).getSheetByName(SHEET_PARAMS).getRange(2,i+1);}
                           return null;})();


/*/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
CONSTANTES : table des ressources TEXT
       - A compléter sur l'ensemble des messages text
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\*/

var SUBJECT ='Dossier %%client%%-%%ao%% pour le %%date%%  - %%unit%%';
var SUBJECT_DECISION_GONOGO ='%%result%%- %%client%%-%%ao%% pour le %%date%%  - %%unit%%';
var SUBJECT_ADMIN='[AVV] Dossiers sans RP au %date%';


/*/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
Fonctions Générales : résolution emailing

/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\*/
var EMAIL_CC=(function () {var params = SpreadsheetApp.openById(CLSID).getSheetByName(SHEET_PARAMS).getDataRange().getValues();
                           for( var i=0;i< params[0].length;i++){if(params[0][i]==COLUMN_CC ) return params[1][i];}
                           return null;})();
function GetEMAIL_CC(){return EMAIL_CC;}
function GetEMAIL_CONTACT(){return enumParams(COLUMN_CONTACT).join();}
function GetEMAIL_ADMIN(){return enumParams(COLUMN_ADMIN).join();}
function GetEMAIL_NOTIFICATION(){return enumParams(COLUMN_NOTIFICATION).join();}
function GetEMAIL_DIFFUSION_GO_CC(){return enumParams(COLUMN_NOTIFICATION_GO_CC).join();}
function GetEMAIL_DIFFUSION_NOGO_CC(){return enumParams(COLUMN_NOTIFICATION_NOGO_CC).join();}

function GetEMAIL_DIFFUSION_NOGO_TO(){return enumParams(COLUMN_NOTIFICATION_NOGO_TO).join();}




function GetIMAGE_URL(id){return 'https://drive.google.com/uc?export=view&id='+id;}
function IsIntegration(){return !(__PROD__);}
var URL_EXEC  = 
  (function () {
  return IsIntegration()?'https://script.google.com/a/macros/sqli.com/s/AKfycbykrG6y4NPDzPIJfRz_EN29THb7eYU2gwFxPnS1Wek/dev'
  :'https://script.google.com/a/macros/sqli.com/s/AKfycbxCQqx63aOJGOBAb6-uACwa9P_ogZxhO2XioEdlRo6UnKpXoiE/exec';})();
  

/*/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
CONSTANTES : MODE DEBUGGING

/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\*/

var __DEBUG__ = true;


//DriveApp.getFolderById(scriptProperties.getProperty('ID_FOLDER_DRIVE_TEMPLATE'));

/*/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
https://sites.google.com/site/scriptsexamples/custom-methods/betterlog
"OFF","SEVERE","WARNING","INFO","CONFIG","FINE","FINER","FINEST" or "ALL" at runtime without editing code.
/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\*/
var Logger = BetterLog.setLevel(PropertiesService.getScriptProperties().getProperty('ID_LOG_LEVEL')).useSpreadsheet(CLSID);

/*function initLevelLog() {var level = PropertiesService.getScriptProperties().getProperty('ID_LOG_LEVEL');
  var value = ["OFF","SEVERE","WARNING","INFO","CONFIG","FINE","FINER","FINEST","ALL"].lastIndexOf(level)!=-1?level:"SEVERE";
  Logger.setLevel(value);
                             
};*/
                             
function treatmentException_(e){
  e = (typeof e === 'string')? new Error(e) :e;
  Logger.warning('(login:%-20s) %s:%s (file %s line %s)',Session.getActiveUser().getEmail()||'',e.name||'',e.message||'',e.fileName||'',e.lineNumber||'');
  throw(e);
};

                             
                             

function SHEET_AO (){
   
   this.sheet =  SpreadsheetApp.openById(CLSID)
                           .getSheetByName(SHEET_O_AVV),
   this.sheet_form = SpreadsheetApp.openById(CLSID)
                           .getSheetByName(SHEET_O_HISTO);
   this.state = GetRowParams(COLUMN_STATE);
   this.rowA=null;this.rowB=null;
}

SHEET_AO.prototype = {
  
  Info_A:function(reference){
    var row, data =getRowsData(this.sheet, this.sheet.getRange(2, 1, this.sheet.getMaxRows() - 1, MAX_COL));
     
      
      for(var i=0;i<data.length;i++){
        row = data[i];
        row.rowNumber = i+2;
        if(reference==row.identifiant) { this.rowA = row;return row;}
        
      }
    return null;
  
  },
  
  Info_B:function(reference){
    var row, data =getRowsData(this.sheet_form, this.sheet_form.getRange(2, 1, this.sheet_form.getMaxRows() - 1, MAX_COL));
      
      
      for(var i=0;i<data.length;i++){
        row = data[i];
        row.rowNumber = i+2;
        if(reference==row.identifiant ) {this.rowB = row;return row;}
        
      }
    return null;
  
  },
  
  getA:function(name){
    return (this.rowA==null)?null:this.rowA[name];
  },
  
  getB:function(name){
    return (this.rowB==null)?null:this.rowB[name];
  },
  
  
}
 


  

function global_testing() {
  var sheet_objet = new SHEET_AO();
  var ref = 'I1433242948271';
  
  var out = sheet_objet.Info_A(ref);
  var out_json;
  Logger.log('out: %s',JSON.stringify(out,null,'\t'));
  //sheet_objet.Info_B(ref);
  Logger.log(GetEMAIL_ADMIN());
  
  
}

