// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const SampleNamespace = {};

(function (SampleNamespace) {
    // The max number of web workers to be created
    const g_maxWebWorkers = 4;

    // The array of web workers
    const g_webworkers = [];

    // Next job id
    let g_nextJobId = 0;

    // The promise info for the job. It stores the {resolve: resolve, reject: reject} information for the job.
    const g_jobIdToPromiseInfoMap = {};

    function getOrCreateWebWorker(jobId) {
        const index = jobId % g_maxWebWorkers;
        if (g_webworkers[index]) {
            return g_webworkers[index];
        }

        // create a new web worker
        const webWorker = new Worker("functions-worker.js");
        webWorker.addEventListener('message', function (event) {
            let jobResult = event.data;
            if (typeof (jobResult) == "string") {
                jobResult = JSON.parse(jobResult);
            }

            if (typeof (jobResult.jobId) == "number") {
                const jobId = jobResult.jobId;
                // get the promise info associated with the job id
                const promiseInfo = g_jobIdToPromiseInfoMap[jobId];
                if (promiseInfo) {
                    if (jobResult.error) {
                        // The web worker returned an error
                        promiseInfo.reject(new Error());
                    }
                    else {
                        // The web worker returned a result
                        promiseInfo.resolve(jobResult.result);
                    }
                    delete g_jobIdToPromiseInfoMap[jobId];
                }
            }
        });

        g_webworkers[index] = webWorker;
        return webWorker;
    }

    // Post a job to the web worker to do the calculation
    function dispatchCalculationJob(functionName, parameters) {
        const jobId = g_nextJobId++;
        return new Promise(function (resolve, reject) {
            // store the promise information.
            g_jobIdToPromiseInfoMap[jobId] = { resolve: resolve, reject: reject };
            const worker = getOrCreateWebWorker(jobId);
            worker.postMessage({
                jobId: jobId,
                name: functionName,
                parameters: parameters
            });
        });
    }

    SampleNamespace.dispatchCalculationJob = dispatchCalculationJob;
})(SampleNamespace);


CustomFunctions.associate("TEST", function (n) {
    return SampleNamespace.dispatchCalculationJob("TEST", [n]);
});

CustomFunctions.associate("TEST_PROMISE", function (n) {
    return SampleNamespace.dispatchCalculationJob("TEST_PROMISE", [n]);
});

CustomFunctions.associate("TEST_ERROR", function (n) {
    return SampleNamespace.dispatchCalculationJob("TEST_ERROR", [n]);
});

CustomFunctions.associate("TEST_ERROR_PROMISE", function (n) {
    return SampleNamespace.dispatchCalculationJob("TEST_ERROR_PROMISE", [n]);
});


// This function will show what happens when calculations are run on the main UI thread.
// The task pane will be blocked until this method completes.
CustomFunctions.associate("TEST_UI_THREAD", function (n) {
    let ret = 0;
    for (let i = 0; i < n; i++) {
        ret += Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(50))))))))));
        for (let l = 0; l < n; l++) {
            ret -= Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(50))))))))));
        }
    }
    return ret;
});

/**
 * @customfunction
 * @param {any} key Key in the key-value pair to set 
 * @param {any} value Value in the key-value pair to set 
*/
function SetValue(key, value) {
    return OfficeRuntime.storage.setItem(key, value).then(function (result) {
        return "Success: Item with key '" + key + "' saved to storage.";
    }, function (error) {
        return "Error: Unable to save item with key '" + key + "' to storage. " + error;
    });
}
CustomFunctions.associate("SETVALUE", SetValue);

/**
 * @customfunction
 * @param {any} key Key the value of which to get
*/
function GetValue(key) {
    return OfficeRuntime.storage.getItem(key);
}
CustomFunctions.associate("GETVALUE", GetValue);

/**
 * Returns data from a web service on the Internet or Intranet
 * @customfunction
 * @param {string} url
 * @return {string} data from a web service on the Internet or Intranet
*/
async function WebService(url) {
    const response = await fetch(url);
    if (!response.ok) {
        throw new Error(response.statusText);
    }
    return response.text();
}
CustomFunctions.associate("WEBSERVICE", WebService);

// /**
//  * Returns specific data from XML content by using the specified xpath
//  * @customfunction
//  * @param {string} xml
//  * @param {string} xpath
//  * @return {string[][]} specific data from XML content by using the specified xpath
// */
// async function FilterXml(xml, xpath) {
//     let doc = new DOMParser();
//     let dom = doc.parseFromString(xml, "text/xml");
//     let query = dom.evaluate(xpath, dom, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
//     let results = [];
//     for (let i = 0, length = query.snapshotLength; i < length; ++i) {
//         results.push([query.snapshotItem(i).textContent]);
//     }
//     return results;
// }
// CustomFunctions.associate("FILTERXML", FilterXml);

// /**
//  * Returns a URL-encoded string
//  * @customfunction
//  * @param {string[][]} Text is a string to be URL encoded
//  * @return {Promise<string[][]>} a URL - encoded strings
// */
// async function EncodeUrl(Text) {
//     return Text.map((i) => i.map((text) => encodeURIComponent(text)));
// }
// CustomFunctions.associate("ENCODEURL", EncodeUrl);

/**
 * Returns a completion for the message
 * @customfunction
 * @param {string} message The message to generate chat completions for.
 * @param {number} [temperature] The sampling temperature between 0.0 and 2 Higher values like 0.8 will make the output more random, while lower values like 0.2 will make it more focused and deterministic (default: 0.5).
 * @param {number} [max_tokens] The maximum number of tokens to generate in the chat completion (default: 256).
 * @param {string} [model] The model to use (default: gpt-4).
 * @param {number} [width] The maximum number of characters per line (default: 68).
 * @param {string} [apiKey] OpenAI api key.
 * @return {string} Completion for the message.
*/
async function ChatGpt(
    message,
    temperature,
    max_tokens,
    model,
    width,
    apiKey
) {
    // Set the HTTP headers
    if (apiKey == null) apiKey = GetValue("key");
    const headers = new Headers();
    headers.append("Content-Type", "application/json");
    headers.append("Authorization", `Bearer ${apiKey}`);

    // Set the HTTP body
    if (model == null) model = "gpt-4"; //gpt-3.5-turbo
    const prompt = [
        { role: "system", content: "You are a helpful assistant." },
        { role: "user", content: message }
    ];
    if (max_tokens < 1 || max_tokens == null || max_tokens > 4000) max_tokens = 256;
    if (temperature < 0 || temperature == null || temperature > 2) temperature = 0.5;
    const body = JSON.stringify({
        model: model,
        messages: prompt,
        max_tokens: max_tokens,
        temperature: temperature
    });

    // Fetch the request
    const response = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: headers,
        body: body
    });

    // Parse the response as JSON
    const json = await response.json();

    if (width == 0 || width == null) width = 68;
    return json.choices[0].message["content"].replace(
        new RegExp(`(?![^\\n]{1,${width}}$)([^\\n]{1,${width}})\\s`, "g"),
        "$1\n"
    );
}
CustomFunctions.associate("CHATGPT", ChatGpt);