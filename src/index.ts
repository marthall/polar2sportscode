import * as XLSX from "xlsx";
import builder from "xmlbuilder";

const onLoad = (fileEvent: Event) => {
  //@ts-ignore
  var files = fileEvent.target.files,
    f = files[0];

  const reader = new FileReader();
  reader.onload = function (e) {
    //@ts-ignore
    var data = new Uint8Array(e.target.result);
    var workbook = XLSX.read(data, { type: "array", cellDates: true });

    const firstPageName = workbook.SheetNames[0];
    const firstPage = workbook.Sheets[firstPageName];
    const json_data: any[][] = XLSX.utils.sheet_to_json(firstPage, {
      header: 1,
    });

    // var cell_address = { c: 0, r: 100 };
    // var cell_ref = XLSX.utils.encode_cell(cell_address);
    //
    // console.log(firstPage[cell_ref]);
    // console.log(firstPage["B100"]);
    console.log(json_data);
    writeXml(json_data);
  };

  reader.readAsArrayBuffer(f);
};

const writeXml = (json_data: any[][]) => {
  const file = builder.create("file", { headless: true });

  const startTime: Date = json_data[1][1];

  const yyyy = startTime.getFullYear();
  const MM = startTime.getMonth().toString().padStart(2, "0");
  const dd = startTime.getDate().toString().padStart(2, "0");
  const HH = startTime.getHours().toString().padStart(2, "0");
  const mm = startTime.getMinutes().toString().padStart(2, "0");
  const ss = startTime.getSeconds().toString().padStart(2, "0");
  const SS = startTime
    .getMilliseconds()
    .toString()
    .padStart(2, "0")
    .slice(0, 2);

  file
    .ele("SESSION_INFO")
    .ele("start_time", {}, `${yyyy}-${MM}-${dd} ${HH}:${mm}:${ss}${SS} +0000`);

  const allInstances = file.ele("ALL_INSTANCES");

  for (let i = 10; i < 10000; i += 10) {
    const id = i / 10;
    const instance = allInstances.ele("instance");
    instance.ele("id", undefined, id);
    instance.ele("code", undefined, "Kevin");

    const distance = instance.ele("label");
    distance.ele("group", undefined, "Distance(m)");
    distance.ele("text", undefined, json_data[i][4].toFixed(3));

    if (json_data[i][2]) {
      const hr = instance.ele("label");
      hr.ele("group", undefined, "HR [bmp]");
      hr.ele("text", undefined, json_data[i][2]);
    }

    const start: Date = json_data[i - 10][1];
    const end: Date = json_data[i][1];
    // @ts-ignore
    const startSeconds = (start - startTime) / 1000;
    // @ts-ignore
    const endSeconds = (end - startTime) / 1000;

    const startEle = instance.ele("start", undefined, startSeconds.toString());
    const endEle = instance.ele("end", undefined, endSeconds.toString());
  }

  console.log(file.end({ pretty: true }));

  download("filename.xml", file.end({ pretty: true }));
};

const download = (filename: string, text: string) => {
  var element = document.createElement("a");
  element.setAttribute(
    "href",
    "data:text/plain;charset=utf-8," + encodeURIComponent(text)
  );
  element.setAttribute("download", filename);

  element.style.display = "none";
  document.body.appendChild(element);

  element.click();

  document.body.removeChild(element);
};

window.onload = () => {
  console.log("Loaded");
  document.getElementById("file-input").addEventListener("change", onLoad);
};
