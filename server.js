const http = require('http');
const url = require('url');

const start = (route, handle) => {
	const onRequest = (request, response) => {
		const pathName = url.parse(request.url).pathname;
		console.log('Request for ' + pathName + ' received.');
		route(handle, pathName, response, request);
	}

	const port = 8000;
	http.createServer(onRequest).listen(port);
	console.log('Server has started. Listening on port: ' + port + '...');
}

exports.start = start;
