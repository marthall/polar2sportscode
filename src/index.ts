import * as XLSX from "xlsx";
import builder from "xmlbuilder";

const STEPS_PER_ENTRY = 10;

const TIME = 1
const HR = 2
const SPEED = 3
const DISTANCE = 4
const ACCELERATION = 5
const CADENCE = 6


const onLoad = (fileEvent: Event) => {
  //@ts-ignore
  var files = fileEvent.target.files,
    f : File = files[0];

  const reader = new FileReader();
  reader.onload = function (e) {
    //@ts-ignore
    var data = new Uint8Array(e.target.result);
    var workbook = XLSX.read(data, { type: "array", cellDates: true });

    let rows : any[][] = []

    for (let sheetIndex = 0; sheetIndex < workbook.SheetNames.length; sheetIndex++) {
      const sheetName = workbook.SheetNames[sheetIndex];
      const sheet = workbook.Sheets[sheetName];
      const data: any[][] = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
      });

      rows.push(...data.slice(1))
    }

    console.log(rows);
    writeXml(f.name, rows);
  };

  reader.readAsArrayBuffer(f);
};

const writeXml = (filename: string, json_data: any[][]) => {
  const file = builder.create("file", { headless: true });

  const startTime: Date = json_data[1][TIME];

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

  for (let i = 1; i < (10000 - STEPS_PER_ENTRY) / STEPS_PER_ENTRY; i++) {
    const instance = allInstances.ele("instance");

    const firstIndex = i * STEPS_PER_ENTRY;
    const lastIndex = i * STEPS_PER_ENTRY + STEPS_PER_ENTRY;

    const start: Date = json_data[firstIndex][TIME];
    const end: Date = json_data[lastIndex][TIME];
    // @ts-ignore
    const startSeconds = (start - startTime) / 1000;
    // @ts-ignore
    const endSeconds = (end - startTime) / 1000;

    instance.ele("start", undefined, startSeconds.toFixed(2));
    instance.ele("end", undefined, endSeconds.toFixed(2));

    instance.ele("id", undefined, i);
    instance.ele("code", undefined, filename);

    const distanceElement = instance.ele("label");
    const distance = json_data[lastIndex][DISTANCE] - json_data[firstIndex][DISTANCE];
    distanceElement.ele("group", undefined, "Distance(m)");
    distanceElement.ele("text", undefined, distance.toFixed(3));

    const hrs: number[] = json_data
      .slice(firstIndex, lastIndex)
      .map((row) => row[HR])
      .filter((value) => !!value);
    if (hrs.length > 0) {
      const maxHRelement = instance.ele("label");
      maxHRelement.ele("group", undefined, "HR (max) [bmp]");
      maxHRelement.ele("text", undefined, Math.max(...hrs).toString());

      const minHRelement = instance.ele("label");
      minHRelement.ele("group", undefined, "HR (min) [bmp]");
      minHRelement.ele("text", undefined, Math.min(...hrs).toString());
    }

    const speeds: number[] = json_data
        .slice(firstIndex, lastIndex)
        .map((row) => row[SPEED])
        .filter((value) => !!value);
    if (speeds.length > 0) {
      const maxSpeedElement = instance.ele("label");
      maxSpeedElement.ele("group", undefined, "Speed (max) [km/h]");
      maxSpeedElement.ele("text", undefined, Math.max(...speeds).toFixed(2));

      const avgSpeedElement = instance.ele("label");
      const speedSum = speeds.reduce((a, b) => a + b, 0);
      const speedAvg = (speedSum / speeds.length) || 0;
      avgSpeedElement.ele("group", undefined, "Speed (avg) [km/h]");
      avgSpeedElement.ele("text", undefined, speedAvg.toFixed(2));
    }

    const accelerations: number[] = json_data
        .slice(firstIndex, lastIndex)
        .map((row) => row[ACCELERATION])
        .filter((value) => !!value);
    if (accelerations.length > 0) {
      const maxAccelerationElement = instance.ele("label");
      maxAccelerationElement.ele("group", undefined, "Acceleration (max) [m/s2]");
      maxAccelerationElement.ele("text", undefined, Math.max(...accelerations).toFixed(2));

      const minAccelerationElement = instance.ele("label");
      minAccelerationElement.ele("group", undefined, "Acceleration (min) [m/s2]");
      minAccelerationElement.ele("text", undefined, Math.min(...accelerations).toFixed(2));
    }
  }

  console.log(file.end({ pretty: true }));


  const outfileName = filename.split(".").slice(0, -1).join(".") + ".xml"
  download(outfileName, file.end({ pretty: true }));
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
