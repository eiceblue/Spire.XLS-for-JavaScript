<template>
  <div class="app-description">
    <h1>Welcome to the Vue Demo</h1>
    <p>
      We are excited to introduce our Vue application, showcasing the powerful capabilities of Spire.XLS for JavaScript in handling Excel files. 
      This project seamlessly integrates the functionalities of Spire.XLS for JavaScript through modern web application frameworks. 
    </p>
    <p class="title">Getting Started</p>
    <ol>
      <li>Copy the two files (<span class="getstart">Spire.Xls.Base.js & Spire.Xls.Base.wasm</span>) from <span class="getstart">lib</span> directory to the <span class="getstart">public</span> directory in Vue project.</li>
      <li>Expand Spire.XLS from the left menu, select a detailed feature and click the "Start" button to execute.</li>
    </ol>
    <p class="title">Product Introduction</p>
    <p>
      Spire.XLS for JavaScript is a powerful Excel development library that allows developers to create, read, modify, and convert Excel files directly within any JavaScript environment. 
      With Spire.XLS for JavaScript, developers can efficiently handle Excel programming tasks—whether generating new documents from templates, updating existing documents, or performing file conversions.
    </p>
    <p class="title">Technical Architecture</p>
    <p>
      Built using WebAssembly technology, Spire.XLS for JavaScript is fully compatible with popular frameworks such as Vue, React, Angular, and plain JavaScript, 
      enabling developers to easily create and distribute their JavaScript applications on the web.
    </p>
    <p class="title">File Support</p>
    <p>
      This versatile JavaScript Excel API operates independently, with no reliance on Microsoft Office Excel. 
      Spire.XLS for JavaScript supports both the legacy Excel 97-2003 format (.xls) and recent versions, including Excel 2007, 2010, 2013, 2016, 2019 (.xlsx, .xlsb, .xlsm), 
      as well as OpenOffice (.ods) format. Compared to custom solutions or Microsoft Automation, Spire.XLS for JavaScript offers superior performance, reliability, and ease of use.
    </p>
    <p class="title">Contact Us</p>
    <ul>
      <li>For any technical questions regarding the product, feel free to contact us: <a href="mailto:support@e-iceblue.com">support@e-iceblue.com</a>.</li>
      <li>For any questions about purchasing the product, please reach out to: <a href="mailto:sales@e-iceblue.com">sales@e-iceblue.com</a>.</li>
      <li>If you can't find the example you need, contact us for a free custom demo: <a href="mailto:support@e-iceblue.com">support@e-iceblue.com</a>.</li>
    </ul>
  </div>
</template>

<script>
export default {
  name: 'AppDescription',
};
</script>

<style scoped>
.app-description {
  padding: 20px;
  font-family: Arial, sans-serif;
  text-align: left;
}
.title {
  text-align: left;
  font-weight: bold;
  color: #2980B9;
  font-size: 17px;
}

h1 {
  color: #1F618D;
}

p {
  line-height: 1.6;
}
.getstart
{
  color: #E67E22 ;
}
ul {
  list-style-type: none;
  padding: 0;
}
ol {
  list-style-type: style type num;;
  padding: 0;
}
b
{
  font-weight: bold
};
ul li{
  margin: 10px 0;
}
</style>