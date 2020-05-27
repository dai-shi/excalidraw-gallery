const fetchData = async (spreadsheetId, gid) => {
  try {
    let sheetName;
    if (gid) {
      const res = await gapi.client.sheets.spreadsheets.get({
        spreadsheetId,
      });
      const sheet = res.result.sheets.find((s) => s.properties.sheetId === Number(gid));
      if (sheet) {
        sheetName = sheet.properties.title;
      }
    }
    const res = await gapi.client.sheets.spreadsheets.values.get({
      spreadsheetId,
      range: (sheetName ? `${sheetName}!` : "") + "A:ZZ",
    });
    const { values } = res.result;
    return values;
  } catch (e) {
    if (e.status === 404) {
      window.alert("Not Found");
      throw e;
    } else if (e.status !== 403) {
      throw e;
    }
  }
  try {
    const auth = gapi.auth2.getAuthInstance();
    await auth.signIn();
    return fetchData(spreadsheetId, gid);
  } catch (e) {
    if (e.error === "popup_blocked_by_browser") {
      window.alert("Popup blocked");
    } else {
      throw e;
    }
  }
};

const main = async () => {
  await gapi.client.init({
    apiKey: "AIzaSyD7XbFJ-EfM5FiAOVWtb38U4bse1vHMuvI",
    discoveryDocs: ["https://sheets.googleapis.com/$discovery/rest?version=v4"],
    clientId: "1075394246632-jds7gs24qr5mkqsjbcu4udhiu5tf0bpu.apps.googleusercontent.com",
    scope: "https://www.googleapis.com/auth/spreadsheets.readonly",
  });
  const container = document.getElementById("container");
  const hash = location.hash.slice(1);
  const searchParams = new URLSearchParams(hash);
  const spreadsheetId = searchParams.get("spreadsheetId");
  const gid = searchParams.get("gid");
  if (!spreadsheetId) {
    container.removeChild(container.firstChild);
    document.getElementById("toolbar").style.display = "block";
    return;
  }
  const values = await fetchData(spreadsheetId, gid);
  if (!values) {
    return;
  }
  const table = document.createElement("table");
  values.forEach((row) => {
    const tr = document.createElement("tr");
    table.appendChild(tr);
    row.forEach((item) => {
      const td = document.createElement("td");
      const match = /https:\/\/excalidraw.com\/#(\S+)/.exec(item);
      if (match) {
        const iframe = document.createElement("iframe");
        iframe.src = `https://dai-shi.github.io/excalidraw-animate/#toolbar=no&autoplay=no&${match[1]}`;
        td.appendChild(iframe);
        const anchor = document.createElement("a");
        anchor.href = item;
        anchor.target = "_blank";
        anchor.rel = "noopener noreferrer";
        anchor.textContent = "Open Excalidraw";
        td.appendChild(anchor);
      } else {
        const span = document.createElement("span");
        span.textContent = item;
        td.appendChild(span);
      }
      tr.appendChild(td);
    });
  });
  container.removeChild(container.firstChild);
  container.appendChild(table);
};

const init = () => {
  gapi.load("client:auth2", main);
};

window.addEventListener("load", init);

window.loadLink = (event) => {
  event.preventDefault();
  const match = /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)[^#&]*(?:[#&]gid=([0-9]+))?/.exec(event.target.link.value);
  if (!match) {
    window.alert("Invalid URL");
    return;
  }
  window.location.hash = `#spreadsheetId=${match[1]}` + (match[2] ? `&gid=${match[2]}` : "");
  window.location.reload();
};
