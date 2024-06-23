function buildUrl_(url, params) {
  var paramString = Object.keys(params).map(function(key) {
    return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
  }).join('&');
  return url + (url.indexOf('?') >= 0 ? '&' : '?') + paramString;
}


function checkIfTriggerExists(eventType, handlerFunction) {
  var triggers = ScriptApp.getProjectTriggers();
  var triggerExists = false;
  triggers.forEach(function (trigger) {
    if(trigger.getEventType() === eventType &&
      trigger.getHandlerFunction() === handlerFunction)
      triggerExists = true;
  });
  return triggerExists;
}
