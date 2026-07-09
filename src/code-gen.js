/**
 * Microsoft Graph OneDrive Explorer - Code Generation Module
 * Translates Graph API requests into fully-functional multi-language code snippets.
 */

export function generateCodeSnippets(method, url, headers = {}, body = '') {
  let cleanBody = body;
  if (typeof body === 'object') {
    cleanBody = JSON.stringify(body, null, 2);
  }

  return {
    curl: generateCurl(method, url, headers, cleanBody),
    javascript: generateJavaScript(method, url, headers, cleanBody),
    powershell: generatePowerShell(method, url, headers, cleanBody),
    csharp: generateCSharp(method, url, headers, cleanBody),
    python: generatePython(method, url, headers, cleanBody)
  };
}

function generateCurl(method, url, headers, body) {
  let snippet = `curl -X ${method} "${url}" \\\n`;
  
  Object.entries(headers).forEach(([key, val]) => {
    // Avoid double spaces, protect quotes
    const cleanVal = String(val).replace(/"/g, '\\"');
    snippet += `  -H "${key}: ${cleanVal}" \\\n`;
  });

  if (body && ['POST', 'PUT', 'PATCH'].includes(method)) {
    const singleLineBody = body.replace(/"/g, '\\"').replace(/\n/g, ' ');
    snippet += `  -d "${singleLineBody}"`;
  } else {
    // Trim trailing backslash and newline
    snippet = snippet.trim().replace(/\\$/, '');
  }

  return snippet;
}

function generateJavaScript(method, url, headers, body) {
  let snippet = `const url = "${url}";\n\n`;
  snippet += `const headers = {\n`;
  Object.entries(headers).forEach(([key, val]) => {
    snippet += `  "${key}": "${String(val).replace(/"/g, '\\"')}",\n`;
  });
  snippet += `};\n\n`;

  const hasBody = body && ['POST', 'PUT', 'PATCH'].includes(method);
  
  snippet += `const options = {\n`;
  snippet += `  method: "${method}",\n`;
  snippet += `  headers: headers,\n`;
  if (hasBody) {
    snippet += `  body: JSON.stringify(${body.trim()})\n`;
  } else {
    // Remove trailing comma of headers
    snippet = snippet.replace(/headers: headers,\n/, 'headers: headers\n');
  }
  snippet += `};\n\n`;

  snippet += `fetch(url, options)\n`;
  snippet += `  .then(res => {\n`;
  snippet += `    if (!res.ok) throw new Error(\`HTTP error! status: \${res.status}\`);\n`;
  snippet += `    return res.json();\n`;
  snippet += `  })\n`;
  snippet += `  .then(data => console.log(data))\n`;
  snippet += `  .catch(err => console.error("Error:", err));`;

  return snippet;
}

function generatePowerShell(method, url, headers, body) {
  let snippet = `$Url = "${url}"\n`;
  snippet += `$Headers = @{\n`;
  Object.entries(headers).forEach(([key, val]) => {
    snippet += `  "${key}" = "${String(val).replace(/"/g, '`"')}"\n`;
  });
  snippet += `}\n\n`;

  const hasBody = body && ['POST', 'PUT', 'PATCH'].includes(method);
  if (hasBody) {
    snippet += `$Body = @'\n${body.trim()}\n'@\n\n`;
  }

  snippet += `$Params = @{\n`;
  snippet += `  Uri = $Url\n`;
  snippet += `  Method = "${method}"\n`;
  snippet += `  Headers = $Headers\n`;
  if (hasBody) {
    snippet += `  Body = $Body\n`;
    snippet += `  ContentType = "application/json"\n`;
  }
  snippet += `}\n\n`;
  snippet += `Invoke-RestMethod @Params`;

  return snippet;
}

function generateCSharp(method, url, headers, body) {
  let snippet = `using System.Net.Http;\n`;
  snippet += `using System.Net.Http.Headers;\n`;
  snippet += `using System.Text;\n`;
  snippet += `using System.Threading.Tasks;\n\n`;
  snippet += `public async Task<string> CallGraphApiAsync()\n`;
  snippet += `{\n`;
  snippet += `    using (var client = new HttpClient())\n`;
  snippet += `    {\n`;
  snippet += `        var request = new HttpRequestMessage(new HttpMethod("${method}"), "${url}");\n\n`;

  Object.entries(headers).forEach(([key, val]) => {
    if (key.toLowerCase() === 'authorization') {
      const authVal = String(val).replace('Bearer ', '');
      snippet += `        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", "${authVal}");\n`;
    } else if (key.toLowerCase() !== 'content-type') {
      snippet += `        request.Headers.TryAddWithoutValidation("${key}", "${String(val).replace(/"/g, '\\"')}");\n`;
    }
  });

  const hasBody = body && ['POST', 'PUT', 'PATCH'].includes(method);
  if (hasBody) {
    snippet += `\n        var jsonBody = @"${body.replace(/"/g, '""')}";\n`;
    snippet += `        request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");\n`;
  }

  snippet += `\n        var response = await client.SendAsync(request);\n`;
  snippet += `        response.EnsureSuccessStatusCode();\n`;
  snippet += `        return await response.Content.ReadAsStringAsync();\n`;
  snippet += `    }\n`;
  snippet += `}`;

  return snippet;
}

function generatePython(method, url, headers, body) {
  let snippet = `import requests\nimport json\n\n`;
  snippet += `url = "${url}"\n\n`;
  snippet += `headers = {\n`;
  Object.entries(headers).forEach(([key, val]) => {
    snippet += `    "${key}": "${String(val).replace(/"/g, '\\"')}",\n`;
  });
  snippet += `}\n\n`;

  const hasBody = body && ['POST', 'PUT', 'PATCH'].includes(method);
  if (hasBody) {
    // Attempt to format Python dictionary if valid JSON
    try {
      const parsed = JSON.parse(body);
      snippet += `payload = ${JSON.stringify(parsed, null, 4)}\n\n`;
      snippet += `response = requests.request("${method}", url, headers=headers, json=payload)\n`;
    } catch {
      snippet += `payload = """${body}"""\n\n`;
      snippet += `response = requests.request("${method}", url, headers=headers, data=payload)\n`;
    }
  } else {
    snippet += `response = requests.request("${method}", url, headers=headers)\n`;
  }

  snippet += `print(response.status_code)\n`;
  snippet += `print(response.text)`;

  return snippet;
}
