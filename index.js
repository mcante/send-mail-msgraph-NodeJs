/*
https://www.npmjs.com/package/@azure/msal-node
https://www.npmjs.com/package/@microsoft/microsoft-graph-client
https://www.appsloveworld.com/nodejs/100/40/microsoft-graph-unhandledpromiserejectionwarning-polyfillnotavailable-library
https://learn.microsoft.com/en-us/answers/questions/908709/nodejs-graph-sendmail
*/

const auth = require('./auth');
const { Client } = require('@microsoft/microsoft-graph-client');

const fs = require('fs');

require('isomorphic-fetch');


async function sendMail() {

  // Token de acceso
  const authResponse = await auth.getToken(auth.tokenRequest);
  
  console.log(authResponse)
  
  // Cliente de Graph
  const client = Client.init({
    authProvider: (done) => {
      done(null, authResponse.accessToken);
    }
  });



  // Cuerpo del correo electrónico
  const email = {
    subject: 'Correo con adjunto MRCG',
    toRecipients: [
      {
        emailAddress: {
          address: 'mcante@gmail.com'
        }
      }
    ],
    body: {
      contentType: 'HTML',
      content: '<h3>Este es un correo con un archivo adjunto.</h3>'
    },
    attachments: [
      {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": "file.pdf",
        "contentBytes": fs.readFileSync('file.pdf', { encoding: 'base64' })
      }
    ]
  };

  // Envío del correo electrónico
  client
    .api('/users/mcante@gmail.com/sendMail')
    .post({ message: email })
    .then((res) => {
      //console.log(res);
      console.log("Ok");
    })
    .catch((err) => {
      console.log(err);
    });



}

sendMail();