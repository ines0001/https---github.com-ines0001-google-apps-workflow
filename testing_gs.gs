function shortenUrl() {
  var url = UrlShortener.Url.insert({
    longUrl: ScriptApp.getService().getUrl()
  });
  Logger.log('Shortened URL is "%s".', url.id);
}

function testing_mail(){
  MailApp.sendEmail('ines0001@gmail.com', 'kjhkjhkjh', 'lkjkjljlkjlkjlkjkljlkjlkjlkj', {noReply:false,});
}
