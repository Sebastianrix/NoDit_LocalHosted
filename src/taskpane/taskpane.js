//var watermark_img;
var watermark2_img;
var canvas;
var canvasWidth;
var canvasHeight;
var img;

// Wait for the Office.js library to be loaded and ready
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log("Office.js is ready, initializing the add-in");


        initializeAddIn();
    } else {
        console.log("This add-in is not designed for the current host:", info.host);
    }
});
/*
function preload() {
    console.log("'preLoad();� Started");
    // Load the image you want to edit
   // watermark_img = "https://localhost:3000/assets/watermark.png";
 //   watermark2_img = "https://localhost:3000/assets/icon-80.png";
    //   img = loadImage("https://localhost:3000/assets/watermark.png");
  
    watermark2_img = loadImage("https://localhost:3000/assets/icon-80.png");
    console.log("'preLoad();� Finished");
}

function setup() {
    console.log("'setup();' run");
    canvasWidth = 526;
    canvasHeight = 785;
    canvas = createCanvas(canvasWidth, canvasHeight);
    console.log("The canvas is: " + canvas);
    const appContainer = document.getElementById("app-container");
    appContainer.appendChild(canvas.elt);
    canvas.elt.style.display = "block";
    canvas.elt.style.margin = "auto";
}*/
function fakeSetupPreload() {
    img = loadImage("https://localhost:3000/assets/watermark.png");
    watermark2_img = loadImage("https://localhost:3000/assets/icon-80.png");

    console.log("'setup();' run");
    canvasWidth = 526;
    canvasHeight = 785;
    canvas = createCanvas(canvasWidth, canvasHeight);
    console.log("The canvas is: " + canvas);
    const appContainer = document.getElementById("app-container");
    appContainer.appendChild(canvas.elt);
    canvas.elt.style.display = "block";
    canvas.elt.style.margin = "auto";
    WriteText();
}
function WriteText() {
    console.log("'WriteText();' run");
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const emailBody = result.value;
            // Ensure image is loaded before generating the canvas
            
                Generate(emailBody);
           
        }
    });
}

function Generate(takenText) {
    console.log("'Generate();' run");
    const myText = takenText;
    const lineBreakRegex = /\r\n|\r|\n|[\u2028\u2029\u00A0-\u00FF\u2000-\u206F\u2E00-\u2E7F\\s\\S]*?(?=(?:\\d+\\.\\s|\*\\s|-\\s|[\w+]\\s|[x]\\s))/g;
    const lines = myText.split(lineBreakRegex);

    console.log("'Generate();' 1");
    const tempCanvas = document.createElement("canvas");
    const tempCtx = tempCanvas.getContext("2d");
    tempCtx.font = "16px Arial";

    let maxWidth = 0;
    let maxHeight = 0;
    for (const line of lines) {
        const lineWidth = tempCtx.measureText(line).width;
        maxWidth = Math.max(maxWidth, lineWidth);
        maxHeight += 20;
    }

    console.log("'Generate();' 2");
    const padding = 20;
    const canvasWidthMail = maxWidth + padding * 2;
    const canvasHeightMail = maxHeight + padding;

    console.log("'Generate();' 3");
    resizeCanvas(canvasWidthMail, canvasHeightMail);

    console.log("'Generate();' 4");
    background(255, 255, 255);

    console.log("'Generate();' 5");
    tint(255, 127);

    console.log("'Generate();' 5.5");

   // watermark2_img.save();
    console.log("'Generate();' 6");
    image(watermark2_img, 0, 0);

    console.log("'Generate();' 7");
    textSize(16);


    console.log("'Generate();' 8");
    fill(0);

    let y = padding;
    for (const line of lines) {
        text(line, padding, y);
        y += 20;
    }

    convertToImage(canvasWidthMail, canvasHeightMail);
}

function convertToImage(canvasWidthMail, canvasHeightMail) {
    console.log("'convertToImage();' run");
    console.log("Canvas.elt is: " + canvas.elt);
    const canvasElement = canvas.elt;
    console.log("'convertToImage();' 2");
    const imageDataUrl = canvasElement.toDataURL();
    console.log("Finished = " + imageDataUrl);
    Office.context.mailbox.item.body.setAsync(
        `<img src="${imageDataUrl}" width="${canvasWidthMail}" height="${canvasHeightMail}">`,
        { coercionType: Office.CoercionType.Html },
        function (result) {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error('Error setting email body: ' + result.error.message);
            } else {
                console.log('Email body set with image.');
            }
        }
    );
}

function initializeAddIn() {
    console.log("Initializing the add-in");
    const helloButton = document.getElementById("helloButton");
    if (helloButton) {
        console.log("Found the 'helloButton' element");
        helloButton.onclick = () => {
            console.log("'helloButton' clicked");
            fakeSetupPreload();
        };
    } else {
        console.error("Could not find the 'helloButton' element");
    }
}
