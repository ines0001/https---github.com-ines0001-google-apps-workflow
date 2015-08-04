function Copy_DriveAPI(source, destination){
  
  // Recopie de l'ensemble des fichiers du répertoire courant
  var tick=0,fileList = source.getFiles();
  while (fileList.hasNext())
  {
    tick++;
    fileList.next().makeCopy(destination);
  }
  
  // Recopie de l'ensemble des folders puis récursivité de la fonction
  var folders = source.getFolders();
  while (folders.hasNext())
  {
    var folder = folders.next();
    
    tick+=Copy_DriveAPI(folder,destination.createFolder(folder.getName()));
  }
  return tick;
}

function Logg(format)
{

  var time=new Date();
  Logger.log('[%s]-%s',time.toTimeString(),format);
  
}

function drive_testing(){
  
  try{
   
   
  }catch(e){Logger.log(e);}
}

//USE ID_FOLDER_DRIVE_PARTAGE property script
function get_drive_all_client(){
  var list=[];
 
  var scriptProperties = PropertiesService.getScriptProperties();
  var source = DriveApp.getFolderById(scriptProperties.getProperty('ID_FOLDER_DRIVE_PARTAGE'));
  
  var folders = source.getFolders();
  while (folders.hasNext())
  {
    var folder = folders.next();
    list.push(folder.getName());
    
  }
  
  
  
  return list; 
}

function get_drive_all_dossier(client){
  var list=[];
 
  var scriptProperties = PropertiesService.getScriptProperties();
  var source = DriveApp.getFolderById(scriptProperties.getProperty('ID_FOLDER_DRIVE_PARTAGE'));
  
  var folders = source.getFoldersByName(client);
  
   while (folders.hasNext())
   {
    source = folders.next();
    var subfolders = source.getFolders();
    while( subfolders.hasNext())
    {
      list.push(subfolders.next().getName());
    }
    break;
   }
  Logger.log('return get_drive_all_dossier');Logger.log(list);
  return list; 
}


function make_copy_cdc2template(ref,Id_dest){
    
    var source = DriveApp.getFolderById(ID_FOLDER_FILES), destination;
    var scriptProperties = PropertiesService.getScriptProperties();
   
    destination = DriveApp.getFolderById(Id_dest); // paramètre appel
    source = source.getFoldersByName(ref); // paramètre appel
    
    if(source.hasNext()){
      source = source.next();}
    else{
      return null;
    }
    
    Logger.log(source.getName());
    var array_branch = scriptProperties.getProperty('ID_FOLDER_DRIVE_CDC').split(',');
    for(var i=0;i<array_branch.length;i++){
      destination = destination.getFoldersByName(array_branch[i]);
      if(destination.hasNext()){
        destination = destination.next();}
      else{
        throw 'Template dossier non trouvé: <b>'+array_branch[i]+'</b>non trouvé';
      }
    }
    
   return Copy_DriveAPI(source, destination);  // return number copy files
}

function make_copy_template(client, folder, editors,ref){
  
  if(__DEBUG__) {Logger.log('make_copy_template %s,%s,%s',client,folder,editors);}
  var scriptProperties = PropertiesService.getScriptProperties();
  
  /*\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
  a scripting exception if the folder does not exist or the user does not have permission to access it
  \/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\*/
  var source = DriveApp.getFolderById(scriptProperties.getProperty('ID_FOLDER_DRIVE_TEMPLATE'));
  
  var folder_partage = DriveApp.getFolderById(scriptProperties.getProperty('ID_FOLDER_DRIVE_PARTAGE'));
  
  var destination = folder_partage.getFoldersByName(client);
  
   if (!destination.hasNext()) {
    
    
    destination = folder_partage.createFolder(client);
    Logger.log('create copy client to :'+destination.getName());
  }else{
   
    destination = destination.next();
    
  }
  
 
  if (!destination.getFoldersByName(folder).hasNext()) {
    
    
    destination = destination.createFolder(folder);
    Logger.log('create copy template to :'+destination.getName());
    //Copy_DriveAPI(source,DriveApp.createFolder("COPY_TEMPLATE"));
  }
  else {
    
    Logger.log('already copy exist');
    // Traitement : envoie d'un mail erreur dossier cible existant
    
    throw 'Le dossier <b>'+folder+'</b> existe déjà...';
  }
 
  /*\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
  Voir si pertinence de checker les permissions et accès aux folders partagés source/Destination
  \/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\*/
  
  var access = source.getSharingAccess();
  
  var yt_privacy;
  // Attempt to preserve permissions from Drive
  switch(access) {
    case DriveApp.Access.PRIVATE:
      yt_privacy = "private";
      break;
    case DriveApp.Access.ANYONE:
      yt_privacy = "public";
      break;
    case DriveApp.Access.ANYONE_WITH_LINK:
      yt_privacy = "unlisted";
      break;
    default:
      yt_privacy = "private";
  }
  if(__DEBUG__) Logger.log("Folder: "+source.getName()+" Access: "+access+" YT privacy: "+yt_privacy);


  /*\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
  REcopie de l'ensemble des accès accordé (liste de mail) ainsi que les niveaux de permission accordés
  \/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\*/
  var i;
  var users = source.getEditors();
  
  for (var i=0;i<users.length;i++){
    
    if(__DEBUG__) Logger.log("Folder: "+destination.getName()+" Edit for "+users[i].getEmail()); 
    //destination.addEditor(users[i]);   
  }

  
  users = source.getViewers();
  
  for (var i=0;i<users.length;i++)
  {
    if(__DEBUG__) Logger.log("Folder: "+destination.getName()+" Viewer for "+users[i].getEmail()); 
    //destination.addViewer(users[i]);
  }
  
  //destination.setOwner(source.getOwner());
  //if(__DEBUG__) Logger.log("Folder: "+source.getName()+" Owner:" +source.getOwner().getEmail());
   
  // Ajout des droits pour le demandeur
  if(editors!==undefined) destination.addEditors(editors);
  
 
  Copy_DriveAPI(source,destination);
  
  
  var objet = new SHEET_AO();
  var row = objet.Info_A(ref);
  var objet_log= new OLogger(row.log);
  
  objet.sheet.getRange(row.rowNumber,COLUMN_SHEET_LOG).setValue(objet_log.pushSession('Génération arborescence AVV','INFO',{
                                                                                       Dossier_AVV:'<a target="_blank" href="'+destination.getUrl()+'">'+destination.getName()+'</a>',
                                                                                       Partage:editors,}))
              
  
  return destination.getId();
}

