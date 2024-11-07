/******************************************************
* Module Name: SmartEmail Manager
* Function: provide manager to process smart email task
* Author: Tian, ShaoQin
* Version: 1.00
* Creation Date: Mar, 2016
*******************************************************/

var JSON = {
    "stringify": rteJSONStringify,
    "parse": rteJSONParse
};

var DEFAULT_INBOX = "INBOX";

var Class = lib.smis_Prototype.getClass();

var getLastReceivedTimeForEmail = lib.SmartEmailUtil.getLastReceivedTimeForEmail;

var convertString2Time = lib.SmartEmailUtil.convertString2Time;

var convertTime2String = lib.SmartEmailUtil.convertTime2String;

var record_msgId = lib.SmartEmailUtil.record_msgId;

var isPOP3orPOP3s = lib.SmartEmailUtil.isPOP3orPOP3s;

var isIMAPorIMAPs = lib.SmartEmailUtil.isIMAPorIMAPs;

var isEWS = lib.SmartEmailUtil.isEWS;

var isIMAPs = lib.SmartEmailUtil.isIMAPs;

var trim = lib.SmartEmailUtil.trim;

var containsToken = lib.SmartEmailUtil.containsToken;
var captureToken = lib.SmartEmailUtil.captureToken;
var removeToken = lib.SmartEmailUtil.removeToken;
var decryptToken = lib.MailUtil.decryptToken;
var parseToken = lib.SmartEmailUtil.parseToken;
var isTokenExpiration = lib.SmartEmailUtil.isTokenExpiration;
var isTokenUsed = lib.SmartEmailUtil.isTokenUsed;

var removeTag = lib.SmartEmailUtil.removeTag;

var isEmptyString = lib.ArrayUtil.isEmptyString;

var contains = lib.ArrayUtil.contains;

var indexOf = lib.ArrayUtil.indexOf;

var ReturnCode = lib.SmartEmailReturnCode.getReturnCode();

var newIncidentId = lib.SmartEmailUtil.newIncidentId;

var sendEmail = lib.SmartEmailUtil.sendEmail;

var isHandleProcessedReturnCode = lib.SmartEmailReturnCode.isHandleProcessedReturnCode;

var isProcessRetryReturnCode = lib.SmartEmailReturnCode.isProcessRetryReturnCode;
var isIgnoredReturnCode = lib.SmartEmailReturnCode.isIgnoredReturnCode;
var ifNeedRecoveredReturnCode = lib.SmartEmailReturnCode.ifNeedRecoveredReturnCode;

var getEmailSubject = lib.MailUtil.getEmailSubject;

var getEmailBody = lib.MailUtil.getEmailBody;

// store connection status
var errorStatus = {};

var SmartEmailManagerClass;
SmartEmailManagerClass = Class.create(lib.smis_Manager.getClass(), {

    initialize: function (configItem) {
        this.configItem = configItem;
        this.logger = this.configItem.getLogger();

        this.intId = this.configItem.intId;

        var adapter1Name = this.configItem.SMAdapter;
        this.SMAdapter = new (lib[adapter1Name].getClass())(this.configItem);
        var adapter2Name = this.configItem.EPAdapter;
        this.EPAdapter = new (lib[adapter2Name].getClass())(this.configItem);
        this.validDomains = this.configItem.getConfigParameterValue("validDomains");
        if (null == this.validDomains) {
            this.validDomains = [];
        } else {
            this.validDomains = this.validDomains.split(/[,;]/);
        }
        this.processedFolder = this.configItem.getConfigParameterValue("processedFolder");
        this.errorFolder = this.configItem.getConfigParameterValue("errorFolder");
        this.removeEmail = this.configItem.getConfigParameterValue("deleteEmail");
        this.protocol = this.configItem.getConfigParameterValue("protocol");
        this.host = this.configItem.getConfigParameterValue("server");
        this.port = this.configItem.getConfigParameterValue("port");
        this.ewsUrl = this.configItem.getConfigParameterValue("EWSUrl");
        this.sharedInbox = this.configItem.getConfigParameterValue("sharedInbox");
        this.mailboxName = this.configItem.getConfigParameterValue("mailboxName");
        this.username = this.configItem.getConfigParameterValue("username");
        this.password = this.configItem.getConfigParameterValue("password");
        this.proxyHost = this.configItem.getConfigParameterValue("proxyHost");
        this.proxyPort = this.configItem.getConfigParameterValue("proxyPort");
        this.proxyUsername = this.configItem.getConfigParameterValue("proxyUsername");
        this.proxyPassword = this.configItem.getConfigParameterValue("proxyPassword");
        this.proxyType = this.configItem.getConfigParameterValue("proxyType");
        this.newEmailUpdate = this.configItem.getConfigParameterValue("newEmailUpdate");
        this.emailArchive = this.configItem.getConfigParameterValue("emailArchive");
        this.mailSizeCheck = this.configItem.getConfigParameterValue("attachProcessing");
        // the value of 'attachSize' from smis is 'kb', times 1024 as 'b'
        var att_size = this.configItem.getConfigParameterValue("attachSize");
        this.maxMailSize = null;
        if (null != att_size) {
            this.maxMailSize = att_size * 1024;
        }
        this.sourceFolderName = trim(this.configItem.getConfigParameterValue("inboxFolder"));
        if (this.sourceFolderName == null || this.sourceFolderName.length == 0) {
            this.sourceFolderName = DEFAULT_INBOX;
        }
        this.serverType = "";

        this.errorCodes = this.configItem.getConfigParameterValue("errorCode").split("~");
        this.errorCases = this.configItem.getConfigParameterValue("errorCases").split("~");
        this.createIncidents = this.configItem.getConfigParameterValue("createIncidents").split("~");
        this.emailAdmins = this.configItem.getConfigParameterValue("emailAdmins").split("~");
        this.emailSenders = this.configItem.getConfigParameterValue("emailSenders").split("~");
        this.emailOthers = this.configItem.getConfigParameterValue("emailOthers").split("~");
        this.errorMessages = this.configItem.getConfigParameterValue("messages").split("~");
        this.adminEmail = this.configItem.getConfigParameterValue("adminEmail");
        this.othersEmail = this.configItem.getConfigParameterValue("othersEmail");

        this.incidentTitle = this.configItem.getConfigParameterValue("title");
        this.incidentAssignmentGroup = this.configItem.getConfigParameterValue("assignmentGroup");
        this.incidentImpact = this.configItem.getConfigParameterValue("impact");
        this.incidentUrgency = this.configItem.getConfigParameterValue("urgency");
        this.incidentCategory = this.configItem.getConfigParameterValue("category");
        this.incidentSubcategory = this.configItem.getConfigParameterValue("subcategory");
        this.incidentArea = this.configItem.getConfigParameterValue("subarea");
        this.incidentAffectedService = this.configItem.getConfigParameterValue("affected.service");
        
        this.useDefaultOperator = this.configItem.getConfigParameterValue("useDefaultOperator");
        this.defaultOperator = this.configItem.getConfigParameterValue("defaultOperator");
        
        // Fetch the parameter for OAuth
        this.enableOAuth = this.configItem.getConfigParameterValue("enableOAuth");
        this.clientID = this.configItem.getConfigParameterValue("clientID");
        this.tenantID = this.configItem.getConfigParameterValue("tenantID");
        this.clientSecret = this.configItem.getConfigParameterValue("clientSecret");
        this.authSite = this.configItem.getConfigParameterValue("authSite");
        this.authScope = this.configItem.getConfigParameterValue("authScope");
        
        lib.SmartEmailUtil.initObjectNumberPrefix();
        lib.SmartEmailUtil.setLogger(this.logger);
    },

initMailReceiver: function (isMailServer) {
    if (this.logger.isDebugEnabled()) {
        this.logger.debug("SmartEmailManagerClass", "initMailReceiver: start");
    }

    if (!errorStatus[this.configItem.name]) {
        errorStatus[this.configItem.name] = {};
    }

    if (this.logger.isDebugEnabled()) {
        this.logger.debug("SmartEmailManagerClass", "errorStatus initialized for " + this.configItem.name);
    }

    if (isEWS(this.protocol)) {
        this.serverType = "EWS";
    } else {
        this.serverType = "Exchange Server";
    }

    if (this.enableOAuth) {
        if (isEWS(this.protocol)) {
            this.serverType += " OAuth";
        }
    }

    if (this.sharedInbox) {
        if (isEWS(this.protocol) || isIMAPs(this.protocol)) {
            this.serverType += " SharedInbox";
        }
    }

    if (this.logger.isDebugEnabled()) {
        this.logger.debug("SmartEmailManagerClass", "Server Type: " + this.serverType);
    }

    var mailOptions = {};
    var optionsJson = "";
    var error;

    // Setting up mailOptions for different server types
    if (this.serverType == "Exchange Server") {
        if (trim(this.protocol) && trim(this.host) && trim(this.port) && trim(this.username) && trim(this.password)) {
            mailOptions = {
                "protocol": this.protocol.toLowerCase(),
                "host": this.host,
                "port": this.port,
                "username": this.username,
                "password": this.password,
                "proxyHost": this.proxyHost,
                "proxyPort": this.proxyPort,
                "proxyUsername": this.proxyUsername,
                "proxyPassword": this.proxyPassword,
                "sourceFolderName": this.sourceFolderName
            };
            if (this.logger.isDebugEnabled()) {
                this.logger.debug("SmartEmailManagerClass", "Mail options set for Exchange Server: " + JSON.stringify(mailOptions));
            }
        } else {
            error = "Please input the mandatory values, such as protocol, host, port, username, password!";
            this.logger.error("SmartEmailManagerClass", error);
            throw new Error(error);
        }
    } else if (this.serverType == "Exchange Server SharedInbox") {
        if (trim(this.protocol) && trim(this.host) && trim(this.port) && trim(this.username) && trim(this.password) && trim(this.mailboxName)) {
            mailOptions = {
                "protocol": this.protocol.toLowerCase(),
                "host": this.host,
                "port": this.port,
                "username": this.username,
                "password": this.password,
                "proxyHost": this.proxyHost,
                "proxyPort": this.proxyPort,
                "proxyUsername": this.proxyUsername,
                "proxyPassword": this.proxyPassword,
                "sourceFolderName": this.sourceFolderName,
                "mailboxName": this.mailboxName
            };
            if (this.logger.isDebugEnabled()) {
                this.logger.debug("SmartEmailManagerClass", "Mail options set for Exchange Server SharedInbox: " + JSON.stringify(mailOptions));
            }
        } else {
            error = "Please input the mandatory values, such as protocol, host, port, username, password, mailboxName!";
            this.logger.error("SmartEmailManagerClass", error);
            throw new Error(error);
        }
    } else if (this.serverType == "EWS") {
        if (trim(this.username) && trim(this.password)) {
            mailOptions = {
                "username": this.username,
                "password": this.password,
                "uri": this.ewsUrl,
                "proxyHost": this.proxyHost,
                "proxyPort": this.proxyPort,
                "proxyUsername": this.proxyUsername,
                "proxyPassword": this.proxyPassword,
                "sourceFolderName": this.sourceFolderName
            };
            if (this.logger.isDebugEnabled()) {
                this.logger.debug("SmartEmailManagerClass", "Mail options set for EWS: " + JSON.stringify(mailOptions));
            }
        } else {
            error = "Please input the mandatory values, such as protocol, username, password!";
            this.logger.error("SmartEmailManagerClass", error);
            throw new Error(error);
        }
    } else if (this.serverType == "EWS SharedInbox") {
        if (trim(this.username) && trim(this.password) && trim(this.mailboxName)) {
            mailOptions = {
                "username": this.username,
                "password": this.password,
                "uri": this.ewsUrl,
                "proxyHost": this.proxyHost,
                "proxyPort": this.proxyPort,
                "proxyUsername": this.proxyUsername,
                "proxyPassword": this.proxyPassword,
                "sourceFolderName": this.sourceFolderName,
                "mailboxName": this.mailboxName
            };
            if (this.logger.isDebugEnabled()) {
                this.logger.debug("SmartEmailManagerClass", "Mail options set for EWS SharedInbox: " + JSON.stringify(mailOptions));
            }
        } else {
            error = "Please input the mandatory values, such as protocol, username, password, mailboxName!";
            this.logger.error("SmartEmailManagerClass", error);
            throw new Error(error);
        }
    } else if (this.serverType == "EWS OAuth") {
        if (trim(this.username) && trim(this.clientID) && trim(this.tenantID) && trim(this.clientSecret) 
            && trim(this.authSite) && trim(this.authScope)) {
            mailOptions = {
                "username": this.username,
                "uri": this.ewsUrl,
                "proxyHost": this.proxyHost,
                "proxyPort": this.proxyPort,
                "proxyUsername": this.proxyUsername,
                "proxyPassword": this.proxyPassword,
                "proxyType": this.proxyType,
                "sourceFolderName": this.sourceFolderName,
                "clientID": this.clientID,
                "tenantID": this.tenantID,
                "clientSecret": this.clientSecret,
                "authSite": this.authSite,
                "authScope": this.authScope
            };
            if (this.logger.isDebugEnabled()) {
                this.logger.debug("SmartEmailManagerClass", "Mail options set for EWS OAuth: " + JSON.stringify(mailOptions));
            }
        } else {
            error = "Please input the mandatory values, such as protocol, username, OAuth information!";
            this.logger.error("SmartEmailManagerClass", error);
            throw new Error(error);
        }
    } else if (this.serverType == "EWS OAuth SharedInbox") {
        if (trim(this.username) && trim(this.mailboxName) && trim(this.clientID) && trim(this.tenantID) 
            && trim(this.clientSecret) && trim(this.authSite) && trim(this.authScope)) {
            mailOptions = {
                "username": this.username,
                "uri": this.ewsUrl,
                "proxyHost": this.proxyHost,
                "proxyPort": this.proxyPort,
                "proxyUsername": this.proxyUsername,
                "proxyPassword": this.proxyPassword,
                "proxyType": this.proxyType,
                "sourceFolderName": this.sourceFolderName,
                "mailboxName": this.mailboxName,
                "clientID": this.clientID,
                "tenantID": this.tenantID,
                "clientSecret": this.clientSecret,
                "authSite": this.authSite,
                "authScope": this.authScope
            };
            if (this.logger.isDebugEnabled()) {
                this.logger.debug("SmartEmailManagerClass", "Mail options set for EWS OAuth SharedInbox: " + JSON.stringify(mailOptions));
            }
        } else {
            error = "Please input the mandatory values, such as protocol, username, mailboxName, OAuth information!";
            this.logger.error("SmartEmailManagerClass", error);
            throw new Error(error);
        }
    }
    
    if (this.logger.isDebugEnabled()) {
        this.logger.debug("SmartEmailManagerClass", "Final Mail options: " + JSON.stringify(mailOptions));
    }
    
    optionsJson = JSON.stringify(mailOptions);
    this.mailReceiver = new MailReceiver(this.serverType);

    try {
        if (isMailServer) {
            if (this.logger.isDebugEnabled()) {
                this.logger.debug('SmartEmailManagerClass', 'Opening connection to mail server...');
            }
            this.mailReceiver.open(optionsJson);
        }
        return ReturnCode.SUCCESS;
    } catch (e) {
        this.logger.error("Error during mailReceiver.open(): " + e.message, e);
        return ReturnCode.CONNECTION_FAILURE;
    }
}
} 
        var object = this.configItem.getConfigParameterValue("object");
        if (this.logger.isDebugEnabled()) {
            this.logger.debug("this.configItem.getConfigParameterValue('object')", object);
        }
        object = lib.SmartEmailUtil.getObjectNameByOption(object); // Convert the object to sm object name

        if (containsToken(mailMessage.body)) {
            var token = captureToken(mailMessage.body);
            task.action = "update";
            token = decryptToken(token);
            if (token) {
                mailMessage.body = removeToken(mailMessage.body);
                var data = parseToken(token);
                task.inRecord['token'] = data;
                task.internalId = data["id"];
                
                task.object = data['file'];
                task.name = data['name'];
                if (data['recipient'].toLowerCase() != sender.toLowerCase()) {
                    return ReturnCode.TOKEN_USURPED;
                }
                
                if (isTokenUsed(data)){
                    if (this.logger.isDebugEnabled()) {
                        this.logger.debug("Pre Validation", "Token been used." + msgId);
                    }
                    return ReturnCode.TOKEN_BEEN_USED;
                }
                
                if (isTokenExpiration(data)) {
                    if (this.logger.isDebugEnabled()) {
                        this.logger.debug("Pre Validation", "Token expired." + msgId);
                    }
                    return ReturnCode.TOKEN_EXPIRED;
                }

                if (this.logger.isDebugEnabled()) {
                    this.logger.debug("Pre Validation", "Outbound Email Token found, going to update ticket.  MSG id: " + msgId);
                }

                task.internalObject = new SCFile(data['file']);
                var keyField = lib.SmartEmailUtil.getUniqueKeyField(data['file']);
                task.inRecord['keyField'] = keyField;
                var query = keyField + "=\"" + data["id"] + "\"";
                var ret = task.internalObject.doSelect(query);
                if (ret != RC_SUCCESS) {
                    if (this.logger.isDebugEnabled()) {
                        this.logger.debug("Pre Validation", "Ticket is not found. MSG id:  " + msgId);
                    }
                    return ReturnCode.RECORD_NOT_FOUND;
                }

            } else {
                return ReturnCode.TOKEN_DECRYPT_FAILED;
            }
        } else {
            task.action = "add";
            
            if (this.newEmailUpdate) {
                var id = lib.SmartEmailUtil.parseUpdateTagID(mailMessage.subject);
                if (id != null) {
                    var obj = lib.SmartEmailUtil.getObjectNameFromID(id);
                    if (obj != null && obj == object) {
                        var file = new SCFile(obj);
                        var keyField = lib.SmartEmailUtil.getUniqueKeyField(obj);
                        var query = keyField + "=\"" + id + "\"";
                        var ret = file.doSelect(query);
                        if (ret == RC_SUCCESS) {
                            task.action = "update";
                            task.internalId = id;
                            task.object = obj;
                            task.internalObject = file;
                            task.inRecord['keyField'] = keyField;
                            task.inRecord['token'] = {} // pseudo token
                            task.inRecord['token']['file'] = obj;
                            task.inRecord['token']['id'] = id;
                            task.inRecord['token']['action'] = "update";
                            if (this.logger.isDebugEnabled()) {
                                this.logger.debug("Pre Validation", "Outbound Email ticket id found, going to update ticket. MSG id:  " + msgId);
                            }
                        } else {
                            if (this.logger.isDebugEnabled()) {
                                this.logger.debug("Pre Validation", "Ticket" + id + " is not found. MSG id:  " + msgId);
                            }
                        }
                    }
                }
            } 


            if (task.action == "add") {
                task.object = object;
                task.internalObject = new SCFile(task.object);
                if (this.logger.isDebugEnabled()) {
                    this.logger.debug("Pre Validation", "No token found, will create ticket.");
                }
            
            }
        }
        
        // remove empty lines from body
        var body = task.inRecord.body;
          while (body && body.length >= 1 && lib.SmartEmailUtil.trim(body[body.length-1]) == '') {
            body.pop();
          }

        return ReturnCode.SUCCESS;
    },

    getAction: function () {
        return this.action;
    },

    getDestObj: function () {
        //return this.obj;
    },

    /**
     * process the task in queue
     *
     **/
    process: function (task) {    
        
        if (this.logger.isDebugEnabled()) {
            this.logger.debug('SmartEmailManagerClass', 'process');
            this.logger.debug("task.inRecord: ", JSON.stringify(task.inRecord));
            this.logger.debug("task.internalObject: ", task.internalObject);
        }

        var resultCode = ReturnCode.SUCCESS;

        var inRecord = task.inRecord;
        
        //lib.smis_TaskManager.updateTask(task);
        if (this.logger.isDebugEnabled()) {
            this.logger.debug("this.SMAdapter.sendRecord(), task", JSON.stringify(task));
        }
        try {
            resultCode = this.SMAdapter.sendRecord(task, task.action);
        } catch (ex) {        
            return ReturnCode.SERVER_ERROR;
        }
        
        // when update email doesn't include valid field and value, skip to save the email as attachment
        if (this.isUpdateTask(task) && task["updateFieldValueList"] && task["updateFieldValueList"].length == 0) {
            return resultCode;
        }

        if (ReturnCode.SUCCESS == resultCode) {
            var filePath = inRecord.filePath;
            if (this.logger.isDebugEnabled()) {
                this.logger.debug("inRecord.filePath", filePath);
            }

            // save .eml as an attachment
            var fileName = task.object;
            if (fileName == 'incidents' || fileName == 'probsummary' || fileName == 'rootcause' || fileName == 'request' || fileName == 'cm3r' || fileName == 'cm3t') {
                var fileId = task.internalId;
                var fileIdName = lib.SmartEmailUtil.getUniqueKeyField(fileName);
                var query = fileIdName + "=\"" + fileId + "\"";
                var scFile = new SCFile(fileName, SCFILE_READONLY);
                var res = scFile.doSelect(query);
                if (res == RC_SUCCESS) {
                    var smattach = new Attachment();
                    var name = inRecord.subject;
                    smattach.type = "application/octet-stream";
                    //QCCR1E145806
                    smattach.name = name.substring(0,lib.SmartEmailConstants.MAX_ATTACHMENT_FILENAME()) + ".eml";
                    smattach.value = readFile(filePath, "b");
                    vars.$attachmentFromEmail = true;
                    var smAttachId = scFile.insertAttachment(smattach);
                    vars.$attachmentFromEmail = false;
                    if(this.emailArchive){
                          var emailRecord = {};
                          emailRecord["user.from"] = task["email.address"];
                          emailRecord["user.to"] = this.username;
                          emailRecord["reference.id"] = fileId;
                          emailRecord["type"] = "inbound";
                          emailRecord["received.time"] = task["receivedUTCDate"];
                          emailRecord["subject"] = task["subject"];
                          emailRecord["uid"] = [];
                          emailRecord["uid"].push(smAttachId.replace(/^cid:/,""));
                                    lib.EmailArchivingUtil.backupEmailIn(emailRecord);
                              }

                } else {
                    return ReturnCode.SERVER_ERROR;
                }
            }
        }
        return resultCode;
    },

    /**
     * post process
     */
    postProcess: function (task, processStatusCode) {
        if (this.logger.isDebugEnabled()) {
            this.logger.debug('SmartEmailManagerClass', 'postProcess');
        }
        
        if (isHandleProcessedReturnCode(processStatusCode) || isIgnoredReturnCode(processStatusCode)) {
            return this.handleProcessedTask(task);
        } else {
            return this.handleFailedTask(task);
        }
    },

    handleProcessedTask: function (task) {
        return this.moveorDeleteEmail(task, this.removeEmail, this.processedFolder);
    },

    handleFailedTask: function (task) {
        return this.moveorDeleteEmail(task, false, this.errorFolder);
    },

    errorHandling: function (errorCode, task) {
        var index = indexOf(this.errorCodes, errorCode);

        if (index < 0) {
            if (this.logger.isDebugEnabled()) {
                this.logger.debug('SmartEmailManagerClass', "Not Configured Error Case happens, Ignore it." + errorCode);
            }
            return;
        }
        var smisErrorMessage = [];
        var emailSubject = "";
        var emailBody = "";
        var emailRecipients = [];

        var refId = "";
        if(task && task["internal.id"])
        {
            refId = task["internal.id"];
        }
        if(this.errorCases.length - 1 >= index){
            emailSubject = this.errorCases[index];
        }
        if (1 == this.emailSenders[index] || 1 == this.emailOthers[index] || 1 == this.emailAdmins[index]) {

            if(this.errorMessages.length -1 >= index){
                  emailBody = this.errorMessages[index];
            }
        }
        var tempSubject = getEmailSubject(emailSubject,null);
        if (this.logger.isDebugEnabled()) {
            this.logger.debug('SmartEmailManagerClass', "Handling SmartEmail ErrorCode: " + errorCode +"[" + tempSubject + "].");
        }

        if (1 == this.createIncidents[index]) {
                  // not create duplication error incident for connection failure
                  if (ifNeedRecoveredReturnCode(errorCode) && this.isIncidentOpened(errorStatus[this.configItem.name]['E'+errorCode])) {
                        smisErrorMessage.push("Incident " + errorStatus[this.configItem.name]['E'+errorCode] + " already been created");
                        if (this.logger.isDebugEnabled()) {
                            this.logger.debug('SmartEmailManagerClass', "Error incident already created and not closed.");
                        }
                  } else {
                        var intID = this.createErrorIncident(task, tempSubject, errorCode);
                        if(intID){
                            smisErrorMessage.push("Incident " + intID + " is created");
                        }
                  }
        }
        
        var tmpSub, tmpBody;
        
        if (1 == this.emailSenders[index]) {
            if (task && task['email.address']) {
                emailRecipients = [];
                emailRecipients.push(task['email.address']);
                tmpSub = getEmailSubject(emailSubject,task['email.address']);
                tmpBody = getEmailBody(emailBody,task['email.address']) + lib.SmartEmailUtil.getOriginatingEmailInfo(task, task['email.address'], true)
                                                                        + lib.SmartEmailUtil.getUpdateFieldValuesEmailInfo(task, task['email.address']);
                sendEmail(emailRecipients, tmpSub, tmpBody,refId);
                smisErrorMessage.push("Email Send to Sender: " + emailRecipients.join(","));
                if (this.logger.isDebugEnabled()) {
                    this.logger.debug('SmartEmailManagerClass', "Email error " + errorCode +" [" + tempSubject + "] to Sender:" + emailRecipients);
                }
            } else {
                if (this.logger.isDebugEnabled()) {
                    this.logger.debug('SmartEmailManagerClass', "No Sender info found, ignore send email to sender.");
                }
            }

        }
        if (1 == this.emailOthers[index]) {
            if (null == this.othersEmail) {
                this.logger.error('SmartEmailManagerClass', "When set Email to Others to Yes, need provide Others Emails.");
            } else {
                emailRecipients = [];
                var tempEmails = this.othersEmail.split("~");
                var i;
                for (i = 0; i < tempEmails.length; i++){
                  var tempEmail = lib.SmartEmailUtil.populateRecipient(tempEmails[i]);
                  if(tempEmail){
                        emailRecipients.push(tempEmail);
                  } else {
                        this.logger.error('SmartEmailManagerClass', tempEmails[i] + "configured in OthersEmail is been ignored , for it is not email address nor SM operator.");
                  }
                }
                if(emailRecipients.length == 0){
                    this.logger.error('SmartEmailManagerClass', "When set Email to Others to Yes, need provide Others Emails.");
                } else {
                  var len = emailRecipients.length;
                  for(i = 0; i < len; i++){
                        var tmpRec = [];
                        tmpRec.push(emailRecipients[i]);
                        var tmpSub = getEmailSubject(emailSubject,emailRecipients[i]);
                        var tmpBody = getEmailBody(emailBody,emailRecipients[i]) + lib.SmartEmailUtil.getOriginatingEmailInfo(task, emailRecipients[i], false)
                                                                                  + lib.SmartEmailUtil.getUpdateFieldValuesEmailInfo(task, emailRecipients[i]);
                        sendEmail(tmpRec, tmpSub, tmpBody,refId);
                        if (this.logger.isDebugEnabled()) {
                            this.logger.debug('SmartEmailManagerClass', "Email error " + errorCode +"[" + tempSubject + "] to Others:" + tmpRec);
                        }
                  }
                  smisErrorMessage.push("Email Send to Others: " + emailRecipients.join(","));
                }
            }
        }
        if (1 == this.emailAdmins[index]) {
            if(null == this.adminEmail){
                this.logger.error('SmartEmailManagerClass', "When set Email to Admin to Yes,  need provide Admin Email.");
            
            } else {
                  var recipient = lib.SmartEmailUtil.populateRecipient(this.adminEmail);
                  if (null == recipient) {
                      this.logger.error('SmartEmailManagerClass', this.adminEmail + "configured in Admin Email is been ignored , for it is not email address nor SM operator.");
                  } else {
                      emailRecipients = [];
                      emailRecipients.push(recipient);
                      var tmpSub = getEmailSubject(emailSubject,recipient);
                      var tmpBody = getEmailBody(emailBody,recipient) + lib.SmartEmailUtil.getOriginatingEmailInfo(task, recipient, false)
                                                                       + lib.SmartEmailUtil.getUpdateFieldValuesEmailInfo(task, recipient);
                      sendEmail(emailRecipients, tmpSub, tmpBody,refId);
                      smisErrorMessage.push("Email Send to Admin: " + emailRecipients.join(","));
                      if (this.logger.isDebugEnabled()) {
                          this.logger.debug('SmartEmailManagerClass', "Email error " + errorCode +"[" + tempSubject + "] to Admin:" + emailRecipients);
                      }
                  }
            }
       }
        
        if (ifNeedRecoveredReturnCode(errorCode)) {
            // remove record from store
            if (this.logger.isDebugEnabled()) {
                this.logger.debug('SmartEmailManagerClass', "Trying to recover task by add a flag at the record from store");
            }
            if(task && task.inRecord.progress){
            	lib.SmartEmailUtil.update_msgId(task.inRecord.msgId, this.protocol, task.inRecord.progress, lib.SmartEmailConstants.SMARTEMAIL_MSGID_NOT_INQUE());
            }
        }
        var tempMess = smisErrorMessage.join(",");
        if(task){
            if(tempMess){
                task.responseMsg = tempMess + " for error: " + errorCode +" [" + tempSubject + "]";
            }else{
                task.responseMsg = "Error " + errorCode +" [" + tempSubject + "] happens without further action.";
            }
        }
    },
    
    successHandling: function (task) {
        // only works for update email
        if (!this.isUpdateTask(task)) {
            return;
        }
        
        var refId = "";
        if(task["internal.id"])
        {
            refId = task["internal.id"];
        }
        
        var address = task['email.address'];
        if (address) {
            var updateFieldValueInfo = lib.SmartEmailUtil.getUpdateFieldValuesEmailInfo(task, address);
            
            var tmpSub = "";
            var tmpBody = "";
            if (updateFieldValueInfo.length>0) { // no valid field and value
                tmpSub = lib.MailUtil.getEmailMessage("ufv_email_subject", "SmartEmail", address);
                tmpBody = lib.MailUtil.getEmailMessage("ufv_email_body", "SmartEmail", address);
            } else {
                tmpSub = lib.MailUtil.getEmailMessage("ufv_email_subject_invalid", "SmartEmail", address);
                tmpBody = lib.MailUtil.getEmailMessage("ufv_email_body_invalid", "SmartEmail", address);
            }
            tmpBody += lib.SmartEmailUtil.getOriginatingEmailInfo(task, address, true) + updateFieldValueInfo;

            sendEmail([address], tmpSub, tmpBody, refId);

            if (this.logger.isDebugEnabled()) {
                this.logger.debug('SmartEmailManagerClass', "Send email about update field value to Sender:" + address);
            }
        } else {
            if (this.logger.isDebugEnabled()) {
                this.logger.debug('SmartEmailManagerClass', "No Sender info found, ignore send email to sender about update field value.");
            }
        }
    },
    
    isUpdateTask: function (task) {
        if (!task || !task["inRecord"] || !task["inRecord"]["token"] || task["action"] != "update" || task["inRecord"]["token"]["action"] != "update") {
            return false;
        }
        
        return true;
    },
    
    createErrorIncident: function(task, tempSubject, errorCode) {
        if (this.logger.isDebugEnabled()) {
            this.logger.debug('SmartEmailManagerClass, this.incidentTitle', this.incidentTitle);
            this.logger.debug('SmartEmailManagerClass, this.incidentAssignmentGroup', this.incidentAssignmentGroup);
            this.logger.debug('SmartEmailManagerClass, this.incidentImpact', this.incidentImpact);
            this.logger.debug('SmartEmailManagerClass, this.incidentUrgency', this.incidentUrgency);
            this.logger.debug('SmartEmailManagerClass, this.incidentCategory', this.incidentCategory);
            this.logger.debug('SmartEmailManagerClass, this.incidentSubcategory', this.incidentSubcategory);
            this.logger.debug('SmartEmailManagerClass, this.incidentArea', this.incidentArea);
            this.logger.debug('SmartEmailManagerClass, this.incidentAffectedService', this.incidentAffectedService);
        }

        var inci = new SCFile("probsummary");

        inci["brief.description"] = this.incidentTitle;
        inci["assignment"] = this.incidentAssignmentGroup;
        inci["initial.impact"] = this.incidentImpact;
        inci["severity"] = this.incidentUrgency;

        inci["category"] = this.incidentCategory;
        inci["subcategory"] = this.incidentSubcategory;
        inci["product.type"] = this.incidentArea;
        inci["affected.item"] = this.incidentAffectedService;
        // description
        inci["action"] = [];
        inci["action"].push(tempSubject);
        inci["action"].push('***********************');
        var server = '';
        if (isEWS(this.protocol)) {
            server = this.ewsUrl;
        } else {
            server = this.host;
        }
        
        var emailAddress = '';
        var operatorName = '';
        var internalId = '';
        var taskId = '';
        if (task) {
            emailAddress = task['email.address'] || '';
            operatorName = task['operator'] || '';
            internalId = task.internalId || '';
            taskId = task.id || '';
        }
        inci["action"].push('Inbound email server: ' + server);
        inci["action"].push('Sender email address: ' + emailAddress);
            inci["action"].push('Sender name: ' + operatorName);
        inci["action"].push('Record ID: ' + internalId);
        var tokenAction = '';
        if (task) {
            if (task.action == 'add') {
                tokenAction = 'add';
            } else {
                if (task.inRecord['token']) {
                    tokenAction = task.inRecord['token']['action'];
                }
            }
        }
        inci["action"].push('Action to be taken: ' + tokenAction);
        inci["action"].push('SMIS task ID: ' + taskId);
        inci["action"].push('Fields list of validation error: ');
        inci["action"].push('***********************');
        
        var ret = inci.doAction('add');
        if(ret == RC_SUCCESS){
            errorStatus[this.configItem.name]['E'+errorCode] = inci["number"];
            if (this.logger.isDebugEnabled()) {
                this.logger.debug('SmartEmailManagerClass', "Incident " + inci["number"] + " created for error" + errorCode +"[" + tempSubject + "]");
            }
            if(task){
                  //task.responseMsg = "Incident " + inci["number"] + " created for error " + errorCode +"[" + tempSubject + "]";
                return inci["number"];
            }
        } else {
        
            if (this.logger.isDebugEnabled()) {
                this.logger.debug('SmartEmailManagerClass', "Create Incident for error " + errorCode +"[" + tempSubject + "] failed.  for " + RCtoString(ret));
            }
        }
        return null;
    },
    
    isIncidentOpened: function(incidentId) {
    	  if (!incidentId) {
          	return false;
          }
          var isIncidentOpened = false;
          var incident = new SCFile("probsummary", SCFILE_READONLY);
          var fields = ["number", "status"];
      
          incident.setFields(fields);
      
          var findIncident = incident.doSelect("number=\"" + incidentId + "\"");
      
          if (findIncident == RC_SUCCESS) {
              if (incident.status != "closed") {
                  isIncidentOpened = true;
              }
          }
      
          return isIncidentOpened;
    },
    
    /**
     * Check email's subject, body is empty and do not has attachment.
     *
     **/
    isEmptyEmail: function(inRecord) {
        if (this.logger.isDebugEnabled()) {
            this.logger.debug('SmartEmailManagerClass', 'Validate Email is empty or not');
        }
        if (inRecord.subject == "" && inRecord.body == "" && !inRecord.hasAttachments) {
            return true;
        } else {
            return false;
        }
    },

    /**
     * Check Sender is from authorized domain
     *
     **/
    isValidDomain: function (sender) {
        var result = false;
        if (this.logger.isDebugEnabled()) {
            this.logger.debug('SmartEmailManagerClass', 'Validate Sender Email Domain');
        }
        if (null == sender) {
            result = false;
        } else {
            var index = sender.indexOf("@");
            var len = sender.length;
            if (index > 0 && index + 1 < len) {
                result = contains(this.validDomains, sender.slice(index + 1, len));
            }
        }
        return result;
    },

    getOperatorNameByContactInfo: function (cont) {
        if (null == cont) {
            return null;
        }

        var ope = cont["operator.id"];
        if (null == ope || "" == ope) {
            if (this.useDefaultOperator && lib.SmartEmailSecurityUtils.isValidOperator(this.defaultOperator)) {
                if (this.logger.isDebugEnabled()) {
                    this.logger.debug('getOperatorNameByContactInfo', "No operator is associated with contact \"" + cont["name"] + "\", default operator \"" + this.defaultOperator + "\" is used");
                }
                return this.defaultOperator;
            }
        
            return null;
        }
        return ope;
    },

    hasNewRight: function (operator, object, right) {
        return lib.SmartEmailSecurityUtils.checkRight(operator, object, right);
    },

    moveorDeleteEmail: function (task, removeEmail, folder) {
        // The option of deleting mail is prior to move mail.
        if (removeEmail) {
            if (this.logger.isDebugEnabled()) {
                this.logger.debug('SmartEmailManagerClass', "Delete email from mail server");
            }
            try {
                this.mailReceiver.deleteMessage(task.inRecord.msgId);
            } catch (e) {
                this.logger.error("Can not delete the mail. Exception: ", e);
                return ReturnCode.CONNECTION_FAILURE;
            }
        } else if (folder) {
            if (!isPOP3orPOP3s(this.protocol.toLowerCase())) {
                if (this.logger.isDebugEnabled()) {
                    this.logger.debug('SmartEmailManagerClass', "Move email to folder: " + folder);
                }
                var jsonParam = {
                    "msgId": task.inRecord.msgId,
                    "folder": folder
                };
                try {
                    this.mailReceiver.moveMessage(JSON.stringify(jsonParam));
                } catch (e) {
                    this.logger.error("Can not move the mail. Exception: ", e);
                    return ReturnCode.CONNECTION_FAILURE;
                }
            } else {
                if (this.logger.isDebugEnabled()) {
                    this.logger.debug('SmartEmailManagerClass', "POP3 or POP3s can not move email.");
                }
            }
        } else {
            if (this.logger.isDebugEnabled()) {
                if (!isPOP3orPOP3s(this.protocol.toLowerCase())) {
                    this.logger.debug('SmartEmailManagerClass', "For processed email, cause the checkbox 'Delete Email when successfully processed' is not selected and 'Processed Email Folder' is not entered. \nFor failed email, cause the 'Error Email Folder' is not entered. \nThe email will not be moved or deleted.");
                } else {
                    this.logger.debug('SmartEmailManagerClass', "The checkbox 'Delete Email when successfully processed' is not selected. The email will not be deleted. ");
                }
            }
        }
        
        return ReturnCode.SUCCESS;
    },

    /**
     *        finalize work
     **/
    finalize: function (task) {
        if (this.logger.isDebugEnabled()) {
            this.logger.debug('SmartEmailManagerClass', 'finalize');
        }
        // delete .eml file which is in the disk.
        var file = task.inRecord.filePath;
        if (this.logger.isDebugEnabled()) {
            this.logger.debug('SmartEmailManagerClass - task.inRecord', JSON.stringify(task.inRecord));
        }
        
        // delete file only from mail server or task status is not waiting (need to retry)
        if (file && (task.inRecord.isMailServer || (task.status != lib.smis_Constants.TASK_STATUS_WAITING()))) {
            try {
                if (this.logger.isDebugEnabled()) {
                    this.logger.debug('SmartEmailManagerClass - try to delete file', file);
                }
                deleteFile(file);
            } catch (e) {
                this.logger.error("Can not delete the temp file. " + file + " Exception: ", e);
            }
        }
        if(task.inRecord.isMailServer){
            if (this.logger.isDebugEnabled()) {
        	    this.logger.debug('SmartEmailManagerClass', 'Close connection of the mail server.');
        	}
        	this.mailReceiver.close();
        }
        delete task.internalObject;
        delete task.inRecord;
        delete task;
    },
    
    enableInstance: function() {
        if (this.logger.isDebugEnabled()) {
            this.logger.debug('SmartEmailManagerClass', "Enabling instance.");
        }
        delete errorStatus[this.configItem.name];
    }

});

function getClass() {
    return SmartEmailManagerClass;
}
