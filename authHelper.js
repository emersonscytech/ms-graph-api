const process = require('process');

let configuration;
try {
	configuration = require("./configuration.json");
} catch(e) {
	console.error("Please create a file named configuration.json in current folder!");
	console.error("Read configuration-template.json.\n");
	process.exit(-1);
}

const credentials = {
	client: configuration.client,
	auth: {
		tokenHost: 'https://login.microsoftonline.com',
		authorizePath: 'common/oauth2/v2.0/authorize',
		tokenPath: 'common/oauth2/v2.0/token'
	}
};

const oauth2 = require('simple-oauth2').create(credentials);

const redirectUri = configuration.redirectUri;

const scopes = [
	'openid',
	'offline_access',
	'User.Read',
	'Mail.Read',
	'Contacts.Read',
	'Calendars.Read'
];

const getAuthUrl = () => {
	const authUrl = oauth2.authorizationCode.authorizeURL({
		redirect_uri: redirectUri,
		scope: scopes.join(' ')
	});
	console.log('Generated auth url: ' + authUrl);
	return authUrl;
}

const getTokenFromCode = (authCode, callback, response) => {
	let token;

	const options = {
		code: authCode,
		redirect_uri: redirectUri,
		scope: scopes.join(' ')
	};

	oauth2.authorizationCode.getToken(options, (error, result) => {
		if (error) {
			callback(response, error, null);
		} else {
			token = oauth2.accessToken.create(result);
			callback(response, null, token);
		}
	});
}

const refreshAccessToken = (refreshToken, callback) => {
	const tokenObj = oauth2.accessToken.create({
		refresh_token: refreshToken
	});
	tokenObj.refresh(callback);
}

exports.refreshAccessToken = refreshAccessToken;
exports.getTokenFromCode = getTokenFromCode;
exports.getAuthUrl = getAuthUrl;
