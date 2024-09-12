<template>
  <span>Click the following button to get the type of hyperlink</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref} from 'vue';

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");


    const startProcessing = async () => {
      wasmModule=window.wasmModule;
      if (wasmModule) {
        // Load font
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

        // Input file
        let excelFileName='HyperlinksSample2.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Load the document
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({fileName: excelFileName});

				//Get the first worksheet
				const sheet = workbook.Worksheets.get(0);
	
				//Iterate all hyperlinks
				const sb = [];
				const hyperlinks = sheet.HyperLinks;
				for(let i=0; i<hyperlinks.Count; i++) {
					const item = hyperlinks.get(i);
					//Get hyperlink address
					const address = item.Address;
					//Get hyperlink type
					const type = item.Type;
					sb.push(`Link address: ${address}`);
					sb.push(`Link type: ${type}`);
					sb.push("");
				}

        //Save result file
        const outputFileName = 'GetHyperLinkType_out.txt';
        FS.writeFile(outputFileName, sb.join("\n"));

        //Dispose
        workbook.Dispose();
		
        // Read the saved file and convert it to Bolb
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray],{type:'text/plain'});

        // Download the result file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl
    };
  }
};
</script>
