var restify = require('restify');
var builder = require('botbuilder');
var $ = require('jquery');
var unirest = require('unirest');
var CRMWebAPI = require('CRMWebAPI'); 
var adal = require('adal-node'); 

//=========================================================
// Global Variables
//=========================================================

//local port: 9000
//emulator url: http://localhost:9000/
//bot url: http://localhost:3987/api/messages
//KB URL Consturction https://kamichel.microsoftcrmportals.com/knowledgebase/article/KA-01056/en-us

//Facebook App ID: 1816303621946249
//Facebook App Secret: ae1affbdbfc00f82d43b04b48c46a325
//FB Access Token: EAAZAz6yQwu4kBAAjYICeHnFXOsNn3oHHalzWpNyZCWhjeHFlKgRI5dpkfroPU2cWggaoC3eB4IMZBAzE71QiV1CMeRd1IWS2DLRosZB9Teu40dpUcyb7vvZB5D4ZBDiOmDrndHZCm6HAHnMa4wCpBthVlPhZA7lZBZCvGiMPX9iYnl1QZDZD


var portalurl_kb = "https://kamichel.microsoftcrmportals.com/knowledgebase/article/";


//=========================================================
// Connect to Azure AD to retrieve Access Token
//=========================================================
// Client ID: 63d68eba-789f-488b-b7ab-6b319fe2124e

var dynamicsApi;
var authorityHostUrl = 'https://login.windows.net/common';
var clientId = '63d68eba-789f-488b-b7ab-6b319fe2124e';
var server = "https://kamichel01.crm.dynamics.com";
var username = "kamichel@kamichel01.onmicrosoft.com";
var pwd = "pass@word1";
var myAccessToken = "";

function authenticateDynamics(fn){ 
    var context = new adal.AuthenticationContext(authorityHostUrl);
    context.acquireTokenWithUsernamePassword(server, username, pwd, clientId, function (err, tokenResponse) {
		if (err) {
			console.log("adalerror = " + JSON.stringify(err));
			console.log("err tokenResponse = " + JSON.stringify(tokenResponse));
		}else {
			myAccessToken = tokenResponse.accessToken;
			fn(tokenResponse.accessToken);
		}
	});
}

//=========================================================
// Connect to Dynamics 365
//=========================================================
// Client ID: 63d68eba-789f-488b-b7ab-6b319fe2124e

authenticateDynamics(function(myAccessToken){
	var apiconfig = { APIUrl: 'https://kamichel01.crm.dynamics.com/api/data/v8.0/', AccessToken: myAccessToken };
	var crmAPI = new CRMWebAPI(apiconfig);	
});

//=========================================================
// Look Up Dynamics Contact (by Policy Number)
//=========================================================
//query reference: https://github.com/davidyack/Xrm.Tools.CRMWebAPI/wiki/Query-Examples
//Get contacts by policy number: contacts?$filter=startswith(new_policynum,%27A%27)&$select=firstname 

function getContact(policynum, fn){

	var apiconfig = { APIUrl: 'https://kamichel01.crm.dynamics.com/api/data/v8.0/', AccessToken: myAccessToken };
	var crmAPI = new CRMWebAPI(apiconfig);	
		   var queryOptions = { Top:1 , 
                                 FormattedValues:true, 
                                 Select:['contactid', 'new_cartype','new_deductible'],
								 Filter:"new_policynum eq '"+policynum+"'"};

            crmAPI.GetList("contacts",queryOptions).then (
                function (response){
					//console.log(response["List"][0]["firstname"]);
					fn(response["List"][0]["contactid"], response["List"][0]["new_cartype"],response["List"][0]["new_deductible"]);

                 }, 
                 function(error){console.log(error)});
}

//=========================================================
// Fetch Dynamics CRM Knowledge Article 
//=========================================================
//Get KB Article Titles: knowledgearticles?$select=title

function getArticle(phrase, fn){

	var apiconfig = { APIUrl: 'https://kamichel01.crm.dynamics.com/api/data/v8.0/', AccessToken: myAccessToken };
	var crmAPI = new CRMWebAPI(apiconfig);	
		   var queryOptions = { Top:1 , 
                                 FormattedValues:true, 
                                 Select:['title', 'description', 'articlepublicnumber'],
								 Filter:"contains(title,'claims process')",
                                 OrderBy:['articlepublicnumber']};

            crmAPI.GetList("knowledgearticles",queryOptions).then (
                function (response){
					var kbPortalLink = portalurl_kb + response["List"][0]['articlepublicnumber'] + "/en-us";
					var formatedTitle = "[" + response["List"][0]['title'] + "]("+ kbPortalLink +")";
					console.log(response["List"][0]['articlepublicnumber'] + formatedTitle);
					
					fn(kbPortalLink, response["List"][0]['title'], response["List"][0]['description']);
                 }, 
                 function(error){console.log(error)});
}

//=========================================================
// Create new Dynamics Case
//=========================================================
function createCase(thisContactID, what, who, details, pic, fn){
	console.log('John_Chan ');
	var apiconfig = { APIUrl: 'https://kamichel01.crm.dynamics.com/api/data/v8.0/', AccessToken: myAccessToken };
	var crmAPI = new CRMWebAPI(apiconfig);	
	console.log('This contact: ' + thisContactID);
	crmAPI
		.Create("incidents", { "title": "Bot Claim", "new_whathappened": what, "new_whowasatfault": who, "description": details, "new_claimimage": pic, "customerid_contact@odata.bind": "contacts("+thisContactID+")" })
			.then(
				function(r){
					console.log('Created: ' + r);

					// Get case ID
					var queryOptions = { Top:1 , 
                                 FormattedValues:true, 
                                 Select:['ticketnumber'],
								 Filter:"incidentid eq " + r,
                                 OrderBy:['ticketnumber']};

            		crmAPI.GetList("incidents",queryOptions).then (
                		function (response){
							console.log('The creased case has ticket number: ' + r);
							fn(response["List"][0]['ticketnumber']);
						});
				}, 
				function(e){
					console.log(e);
				});

}

//=========================================================
// Check Sentiment
//=========================================================
// Sentiment Key: 3327a97ea8a54fee86c66af983c80d28

	function getSentiment(myText, fn){
		
		unirest.post('https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/sentiment?')
		.headers({'Content-Type': 'application/json', 'Ocp-Apim-Subscription-Key': '3327a97ea8a54fee86c66af983c80d28'})
		.send({ "documents": [{"language": "en", "id": "bot", "text": myText}]})
		.end(function (response) {
		  var myScore = response.body['documents'][0]['score'];
		  console.log(myScore);
		  fn(myScore);
		});
	}		

//=========================================================
// Interpret Picture
//=========================================================
//Computer Vision Key: 23e09036b201446ebdc87e9231fc4e07
//data['categories'][0]['name']

	function getPicture(myPic, fn){
		console.log(myPic);
		var myTestPic = "https://www.microsoft.com/dynamics/AsiaPartnerSummit/Images/speakers/John_Chan.jpg";
		//var myTestPic = "http://blog.oxforddictionaries.com/wp-content/uploads/selfie4.jpg";
		unirest.post('https://api.projectoxford.ai/vision/v1/analyses?visualFeatures=Description&')
		.headers({'Content-Type': 'application/json', 'Ocp-Apim-Subscription-Key': '23e09036b201446ebdc87e9231fc4e07'})
		.send({"visualFeatures":"Description","Url": myTestPic})
		.end(function (response) {
		 // console.log(response.body.description.captions[0]['text']);
		  //Check and see if there is a car mentioned in the tags...
		  var imageTags = JSON.stringify(response.body.description.tags);
		  console.log(imageTags);
		  fn(imageTags.includes("car"), response.body.description.captions[0]['text']);
		});
	}	


//=========================================================
// Bot Setup
//=========================================================
//Application ID: 6f8ce5be-b1cd-4de8-b819-60ae5ce7fc0f
//Application Password: cxjFEBqaCu7PDVGqM94Htuq

	// Setup Restify Server
	var server = restify.createServer();
	server.listen(process.env.port || process.env.PORT || 3978, function () {
	   console.log('%s listening to %s', server.name, server.url); 
	});
	  
	// Create chat bot
	var connector = new builder.ChatConnector({
		appId: process.env.MICROSOFT_APP_ID,
		appPassword: process.env.MICROSOFT_APP_PASSWORD
	});
	var bot = new builder.UniversalBot(connector);
	server.post('/api/messages', connector.listen());

//=========================================================
// Connect to LUIS Model for dialog
//=========================================================
// App ID: 7a7701a1-c3a5-40b7-a256-5ee8ff624aaa 
// API Key: ab871b6b06a9477eb8e9e97aa0ffa2cb

	var model = process.env.model || 'https://api.projectoxford.ai/luis/v1/application?id=7a7701a1-c3a5-40b7-a256-5ee8ff624aaa&subscription-key=ab871b6b06a9477eb8e9e97aa0ffa2cb';
	var recognizer = new builder.LuisRecognizer(model);
	var intents = new builder.IntentDialog({ recognizers: [recognizer] });
	bot.dialog('/', intents);

//=========================================================
// Bots Intentions
//=========================================================
	/* 
	 * DEFAULT
	 */
	intents.onDefault([
		function (session, results) {
			session.send("Hey Archie. I am available to help you 24/7, for whatever you may need.");
		}
	]);
	
	intents.matches('Greeting', [
		function (session) {
			session.send('Hi Archie, how can I help you?');
		}
	]);

	intents.matches('Need Help', [
		function (session) {
			session.beginDialog('/validate');
		}
	]);

	intents.matches('Coverage Related', [
		function (session) {
			session.beginDialog('/coverage');
		}
	]);
	
	intents.matches('Deductible', [
		function (session) {
			session.beginDialog('/deductible');
		}
	]);

	intents.matches('Rental', [
		function (session) {
			session.beginDialog('/rental');
		}
	]);
	
//=========================================================
// Bots Dialogs
//=========================================================

/*
 * /validate
 * Validate who we are talking to and retrieve their policy number.
 */

	bot.dialog('/validate', [
		function (session) {
			   session.userData.policyNum = "";
			   builder.Prompts.text(session, 'Alright, I will be happy to help. Can you please tell me your policy number.');
		},
		function (session, results) {
			var policyNumber = results.response;
			//A user might enter more than just the number, so strip the # with regex
			policyNumber = policyNumber.replace(/[^0-9]/g,'');
			//Save the policy number to the session for later reference
			session.userData.policyNum = policyNumber;
			//Confirm policy number is as recorded
			builder.Prompts.choice(session, 'Thanks. I have your policy number as *'+ session.userData.policyNum +'*. Is that correct?',["Yes", "No"]);
		},
		function (session, results) {
			if(results.response.entity == "Yes"){
				session.send('Hang on just a second while I pull up your policy...');
				getContact(session.userData.policyNum, function(contactid, car, deductible){
					session.userData.car = car;
					session.userData.deductible = deductible;
					session.userData.id = contactid;
					session.send('Great, I found you. How can I help with your ' + session.userData.car);
					session.endDialog();
				});
			}else{
				session.beginDialog('/validate');
			}
		}
	]);

/*
 * /coverage
 * Inquiry dialog about coverage
 */

	bot.dialog('/coverage', [
		function (session) {
			session.send('I think I may have a Knowledgebase article that might help you...');
			getArticle('policy',function(kbLink, kbTitle, kbDesc){
					var msg = new builder.Message(session)
						.attachments([
							new builder.HeroCard(session)
								.images([
									builder.CardImage.create(session, "http://www.thatgoodolehandyman.com/wp-content/uploads/2014/10/charlotte-insurance-claim-form.jpg")	
								])
								 .title(kbTitle)
								 .subtitle(kbDesc)
								.tap(builder.CardAction.openUrl(session, kbLink))
						]);
					session.send(msg);
					session.endDialog();
			});
		}
	]);

/*
 * /deductible
 * Inquiry dialog about deductible
 */

	bot.dialog('/deductible', [
		function (session) {
			//TODO - Check to make sure we're validated, if not, send back to dialog /validate
			session.send('You have a $'+ session.userData.deductible +' deductible on policy '+ session.userData.policyNum);
			//TODO - CalcHistogram function wrapped around this:
			builder.Prompts.choice(session, 'Sounds to me like you may need to start a new claim, am I right?',["Yes", "No"]);
		},
		function (session, results) {
			if(results.response.entity == "Yes"){
				session.beginDialog('/newclaim');
			}else{
				session.send('Alright, what else can I do for you?');
				session.endDialog();
			}
		}
	]);

/*
 * /newclaim
 * New claim submission dialog
 */

	bot.dialog('/newclaim', [
		function (session) {
			//TODO - Check to make sure we're validated, if not, send back to dialog /validate
			session.send('I can help you start a new claim. I have just a few questions and then I will take it from there.');
			builder.Prompts.choice(session, 'First, please tell me what happened to your ' + session.userData.car,["Collision", "Breakdown", "Windshield"]);
		},
		function (session, results) {
			if(results.response.entity == "Collision"){
				session.userData.claimq1 = results.response.entity;
				builder.Prompts.choice(session, 'Ok, now who was at fault?',["Self", "Other Driver", "Unknown"]);
			}else{
				session.send('Sorry mate, I cannot help you with a '+ results.response.entity +' yet, sorry!');
				session.endDialog();
			}
		},
		function (session, results) {
			session.userData.claimq2 = results.response.entity;
			builder.Prompts.text(session, 'Can you please give me a brief description of what happened?');
		},
		function (session, results) {
			session.userData.claimq3 = results.response;
			//session.send('Q1 '+ session.userData.claimq1 + 'Q2: '+ session.userData.claimq2 + 'Q3: ' + session.userData.claimq3);
			builder.Prompts.attachment(session, 'Thanks. One last thing, can you send me a picture of your car?');
		},
		function (session, results) {
			//TODO - Do a check of the picture and list out the description. If no car, ask again.
			session.userData.claimq4 = JSON.stringify(results.response[0]['contentUrl']);
			getPicture(results.response[0]['contentUrl'], function(isCar, picDesc){
				if(isCar){
					session.send('Great, that looks like ' + picDesc);
					next();
				}else{
					builder.Prompts.attachment(session, 'I do not see a car in this picture. It looks like '+ picDesc +'. Can you please send me a picture of your car?');
				}
			});
		},
		function (session, results) {
			//createCase(thisContactID, what, who, details, fn){
			session.send('Perfect. Please hold a moment while I start your claim.');
			session.send('Thanks!. Your new claim number is CAS-.');
		//	createCase(session.userData.id, session.userData.claimq1, session.userData.claimq2, session.userData.claimq3, session.userData.claimq3, function(caseNum){
		//		session.send('Thanks! Your new claim number is *'+ caseNum +'* You can review the status of your claim at any time at our [web portal](https://kamichel.microsoftcrmportals.com/).');
		//		session.send('Please let me know if there is anything else I can assist you with today.');
		//		session.endDialog();
		//	});
		},
	]);

/*
 * /rental
 * Inquiry dialog about a rental car
 */

	bot.dialog('/rental', [
		function (session) {
			//TODO Check if he really has rental on his policy
			builder.Prompts.text(session, 'I see that you do not have rental coverage on your policy ' + session.userData.policyNum + '. Would you like me to use Bing to find a local rental car facility for you?');
		},
		function (session, results) {
			var sentScore = 0;
			getSentiment(results.response, function(myScore){
				if (myScore < .5){
					builder.Prompts.choice(session, 'Sorry, I can tell you are getting frustrated. Would you like me to initiate a warm transfer to a live agent?',["Yes", "No"]);
				}else{
					session.send('Great! Here are your results <<RESULTS>>');
				}
			});
		},
		function (session, results) {
			if(results.response.entity == "Yes"){
				session.send('Please wait a moment while I connect you with a live agent. They will review our conversation and will pick up where I left off.');
				session.endDialog();
			}else{
				session.send('Alright, what else can I do for you?');
				session.endDialog();
				req.session.destroy()
			}
		}
	]);