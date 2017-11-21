const url = require('url');
const server = require('./server');
const router = require('./router');
const authHelper = require('./authHelper');
const microsoftGraph = require("@microsoft/microsoft-graph-client");

const getValueFromCookie = (valueName, cookie) => {
	if (cookie.indexOf(valueName) !== -1) {
		const start = cookie.indexOf(valueName) + valueName.length + 1;
		let end = cookie.indexOf(';', start);
		end = end === -1 ? cookie.length : end;
		return cookie.substring(start, end);
	}
}

const getAccessToken = (request, response, callback) => {
	const tokenExpireIn = parseFloat(getValueFromCookie('ms-auth-token-expires', request.headers.cookie));
	const expiration = new Date(tokenExpireIn);

	if (expiration <= new Date()) {
		const refreshToken = getValueFromCookie('ms-auth-refresh-token', request.headers.cookie);

		authHelper.refreshAccessToken(refreshToken, (error, newToken) => {
			if (error) {
				callback(error, null);
			} else if (newToken) {

				const cookies = [
					'ms-auth-token=' + newToken.token.access_token + ';Max-Age=4000',
					'ms-auth-refresh-token=' + newToken.token.refresh_token + ';Max-Age=4000',
					'ms-auth-token-expires=' + newToken.token.expires_at.getTime() + ';Max-Age=4000'
				];

				response.setHeader('Set-Cookie', cookies);
				callback(null, newToken.token.access_token);
			}
		});
	} else {
		const accessToken = getValueFromCookie('ms-auth-token', request.headers.cookie);
		callback(null, accessToken); // Return cached token
	}
}

const calendar = (response, request) => {
	getAccessToken(request, response, (error, token) => {
		const email = getValueFromCookie('ms-auth-email', request.headers.cookie);

		if (token) {
			response.writeHead(200, {
				'Content-Type': 'application/json'
			});

			const client = microsoftGraph.Client.init({
				authProvider: (done) => {
					done(null, token);
				}
			});

			client.api('/me/calendars').header('X-AnchorMailbox', email).get((err, calendars) => {
				const calendar = calendars.value.filter((calendar) => {
					return calendar.name == "api";
				})[0];

				client
				.api(`/me/calendars/${calendar.id}/events`)
				.header('X-AnchorMailbox', email)
				.filter("start/dateTime ge '2018-01-01T00:00:00'")
				.filter("location/displayName eq 'Imbituba'")
				.get((err, res) => {
					console.log(err)
					response.write(JSON.stringify(res));
					response.end();
				});
			});
		} else {
			response.writeHead(200, {
				'Content-Type': 'text/html; charset=UTF-8'
			});
			response.write('<p> No token found in cookie!</p>');
			response.end();
		}
	});
}

const home = (response, request) => {
	response.writeHead(200, {
		'Content-Type': 'text/html'
	});
	response.write('<p>Please <a href="' + authHelper.getAuthUrl() + '">sign in</a> with your Office 365 or Outlook.com account.</p>');
	response.end();
}

const getUserEmail = (token, callback) => {
	const client = microsoftGraph.Client.init({
		authProvider: (done) => {
			done(null, token);
		}
	});

	client.api('/me').get((err, res) => {
		if (err) {
			callback(err, null);
		} else {
			callback(null, res.userPrincipalName);
		}
	});
}

const authorize = (response, request) => {
	const urlParts = url.parse(request.url, true);
	const code = urlParts.query.code;
	authHelper.getTokenFromCode(code, tokenReceived, response);
}

const tokenReceived = (response, error, token) => {
	if (error) {
		response.writeHead(200, {
			'Content-Type': 'text/html'
		});
		response.write('<p>ERROR: ' + error + '</p>');
		response.end();
	} else {
		getUserEmail(token.token.access_token, (error, email) => {
			if (error) {
				response.write('<p>ERROR: ' + error + '</p>');
				response.end();
			} else if (email) {

				const cookies = [
					'ms-auth-token=' + token.token.access_token + ';Max-Age=4000',
					'ms-auth-refresh-token=' + token.token.refresh_token + ';Max-Age=4000',
					'ms-auth-token-expires=' + token.token.expires_at.getTime() + ';Max-Age=4000',
					'ms-auth-email=' + email + ';Max-Age=4000'
				];

				response.setHeader('Set-Cookie', cookies);

				response.writeHead(301, {
					'Location': 'http://localhost:8000/calendar'
				});

				response.end();
			} else {
				console.log("ELSE!", email)
			}
		});
	}
}

const handle = {
	'/': home,
	'/authorize': authorize,
	'/calendar': calendar
};

server.start(router.route, handle);
