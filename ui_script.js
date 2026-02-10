// ui_script.js
const settingList = {
  book: {
    type: "radio",
    label: "文書の種類　　",
    variable: "isBook",
    options: {
      option1: { text: "書籍", value: "book" },
      option2: { text: "論文、レポート", value: "", default: true },
    },
  },
  papersize: {
    type: "select",
    label: "用紙サイズ　　",
    variable: "paperSize",
    options: {
      siroku: { text: "四六版", value: "paper={130mm,188mm}", default: true },
      a0paper: { text: "A0", value: "a0paper" },
      a1paper: { text: "A1", value: "a1paper" },
      a2paper: { text: "A2", value: "a2paper" },
      a3paper: { text: "A3", value: "a3paper" },
      a4paper: { text: "A4", value: "a4paper" },
      a5paper: { text: "A5", value: "a5paper" },
      a6paper: { text: "A6", value: "a6paper" },
      a7paper: { text: "A7", value: "a7paper" },
      a8paper: { text: "A8", value: "a8paper" },
      a9paper: { text: "A9", value: "a9paper" },
      a10paper: { text: "A10", value: "a10paper" },
      b0paper: { text: "B0", value: "b0paper" },
      b1paper: { text: "B1", value: "b1paper" },
      b2paper: { text: "B2", value: "b2paper" },
      b3paper: { text: "B3", value: "b3paper" },
      b4paper: { text: "B4", value: "b4paper" },
      b5paper: { text: "B5", value: "b5paper" },
      b6paper: { text: "B6", value: "b6paper" },
      b7paper: { text: "B7", value: "b7paper" },
      b8paper: { text: "B8", value: "b8paper" },
      b9paper: { text: "B9", value: "b9paper" },
      b10paper: { text: "B10", value: "b10paper" },
    },
  },
  tate: {
    type: "radio",
    label: "文書の方向　　",
    variable: "isTate",
    options: { option1: { text: "縦書き", value: "tate", default: true }, option2: { text: "横書き", value: "" } },
  },
  twocolumn: {
    type: "radio",
    label: "段組み　　　　",
    variable: "isTwoColumn",
    options: {
      option1: { text: "1段組", value: "", default: true },
      option2: { text: "2段組", value: "twocolumn" },
    },
  },
  jafontsize: {
    type: "select",
    label: "本文フォントサイズ（級数）　　",
    variable: "fontSize",
    options: {
      fontsize8q: { text: "8級", value: "8Q" },
      fontsize9q: { text: "9級", value: "9Q" },
      fontsize10q: { text: "10級", value: "10Q" },
      fontsize11q: { text: "11級", value: "11Q" },
      fontsize12q: { text: "12級", value: "12Q" },
      fontsize13q: { text: "13級", value: "13Q" },
      fontsize14q: { text: "14級", value: "14Q" },
      fontsize15q: { text: "15級", value: "15Q" },
      fontsize16q: { text: "16級", value: "16Q" },
      fontsize17q: { text: "17級", value: "17Q" },
      fontsize18q: { text: "18級", value: "18Q" },
      fontsize19q: { text: "19級", value: "19Q" },
      fontsize20q: { text: "20級", value: "20Q" },
      fontsize6pt: { text: "6pt", value: "6pt" },
      fontsize7pt: { text: "7pt", value: "7pt" },
      fontsize8pt: { text: "8pt", value: "8pt" },
      fontsize9pt: { text: "9pt", value: "9pt" },
      fontsize10pt: { text: "10pt", value: "", default: true },
      fontsize11pt: { text: "11pt", value: "11pt" },
      fontsize12pt: { text: "12pt", value: "12pt" },
      fontsize13pt: { text: "13pt", value: "13pt" },
      fontsize14pt: { text: "14pt", value: "14pt" },
      fontsize15pt: { text: "15pt", value: "15pt" },
    },
  },
};
const settingList2 = {
  defaultheadings: {
    type: "select",
    label: ".docxの「見出しレベル1」に対応させるLaTeXの見出しレベル　　",
    variable: "defaultHeadingTag",
    options: {
      option1: { text: "part", value: 0 },
      option2: { text: "chapter", value: 1 },
      option3: { text: "section", value: 2, default: true },
    },
  },
  autonumbering: {
    type: "checkbox",
    label: "見出しに自動で番号を振る（非推奨）",
    text: "",
    variable: "isAutoNumbering",
    default: false,
  },
};
const settingList3 = {
  fullwidthbracket: {
    type: "checkbox",
    label: "括弧・亀甲括弧内の文章を級数下げする（マクロ使用）",
    text: "",
    variable: "isFullWidthBracket",
    default: false,
  },
  useSmallFontSizeInBracket: {
    type: "checkbox",
    label: "括弧・亀甲括弧内の文章を級数下げする（マクロ不使用）",
    text: "",
    variable: "useSmallFontSizeInBracket",
    default: true,
  },
  convertEMFFile: {
    type: "checkbox",
    label: ".emfファイルを変換する（β版）",
    text: "",
    variable: "isConvertEMFFile",
    default: false,
  },
  convertMath: {
    type: "checkbox",
    label: "数式を変換する",
    text: "",
    variable: "isConvertMath",
    default: true,
  },
};

let optionsArray = new Array();

function buildOptions(formElementList, formContainerId) {
  const formContainer = document.getElementById(formContainerId);
  for (var key in formElementList) {
    const formElement = formElementList[key];
    formElement.id = key;
    createFormElement(formContainer, formElement);
  }
  optionsArray.push(formElementList);
  getVariables();
}

function createFormElement(formContainer, formElement) {
  const elementContainer = document.createElement("div");
  formContainer.appendChild(elementContainer);

  switch (formElement.type) {
    case "radio":
      const label_radio = document.createElement("label");
      if (formElement.label !== "") {
        label_radio.textContent = formElement.label;
        elementContainer.appendChild(label_radio);
      }

      for (var key in formElement.options) {
        const option = formElement.options[key];
        const radioButton = document.createElement("input");
        radioButton.type = "radio";
        radioButton.name = formElement.id;
        radioButton.value = option.value;
        radioButton.onchange = getVariables;
        radioButton.id = `${formElement.id}_${key}`;
        if (option.default) radioButton.checked = true;
        elementContainer.appendChild(radioButton);

        const optionLabel = document.createElement("label");
        optionLabel.textContent = option.text;
        optionLabel.setAttribute("for", radioButton.id);
        elementContainer.appendChild(optionLabel);
      }

      break;

    case "select":
      const label_select = document.createElement("label");
      if (formElement.label !== "") {
        label_select.textContent = formElement.label;
        label_select.htmlFor = formElement.id;
        elementContainer.appendChild(label_select);
      }

      const selectbox = document.createElement("select");
      selectbox.id = formElement.id;
      selectbox.onchange = getVariables;
      elementContainer.appendChild(selectbox);

      for (var key in formElement.options) {
        const option = formElement.options[key];
        const option_add = document.createElement("option");
        option_add.value = option.value;
        option_add.textContent = option.text;
        if (option.default) option_add.selected = true;
        selectbox.appendChild(option_add);
      }
      break;

    case "checkbox":
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.text = formElement.text;
      checkbox.id = formElement.id;
      checkbox.onchange = getVariables;
      if (formElement.default) checkbox.checked = true;
      elementContainer.appendChild(checkbox);

      const label_checkbox = document.createElement("label");
      if (formElement.label !== "") {
        label_checkbox.textContent = formElement.label;
        label_checkbox.htmlFor = formElement.id;
        elementContainer.appendChild(label_checkbox);
      }

      break;
    // その他の要素タイプに対する処理
  }
}

function getVariables() {
  for (let i = 0; i < optionsArray.length; i++) {
    for (var key in optionsArray[i]) {
      const formElement = optionsArray[i][key];

      switch (formElement.type) {
        case "radio":
          const radio_options = document.getElementsByName(formElement.id);
          for (let j = 0; j < radio_options.length; j++) {
            if (radio_options.item(j).checked) {
              window[formElement.variable] = radio_options.item(j).value;
            }
          }
          break;
        case "select":
          const select_element = document.getElementById(formElement.id);
          window[formElement.variable] = select_element.value;
          break;
        case "checkbox":
          const checkbox_element = document.getElementById(formElement.id);
          window[formElement.variable] = checkbox_element.checked;
          break;
      }
    }
  }
  return;
}

const template1Options = {
  option1: {
    text: "Yes",
    checked: true,
  },
  option2: {
    value: "Dog",
    text: "Dog",
  },
  option3: {
    values: ["Lion"],
    texts: ["Lion"],
  },
  option4: {
    value: "Dog Food",
    text: "Dog Food",
  },
  option5: {
    text: "Template 1 Text",
  },
};

const template2Options = {
  option1: {
    text: "Yes",
    checked: false,
  },
  option2: {
    value: "Cat",
    text: "Cat",
  },
  option3: {
    values: ["Bird", "Butterfly"],
    texts: ["Bird", "Butterfly"],
  },
  option4: {
    value: "Cat Food",
    text: "Cat Food",
  },
  option5: {
    text: "Template 2 Text",
  },
};

const template3Options = {
  option1: {
    text: "Yes",
    checked: true,
  },
  option2: {
    value: "Wolf",
    text: "Wolf",
  },
  option3: {
    values: ["Lion", "Bird"],
    texts: ["Lion", "Bird"],
  },
  option4: {
    value: "Churu",
    text: "Churu",
  },
  option5: {
    text: "Template 3 Text",
  },
};

function updateOptions() {
  const templateSelect = document.getElementById("template-select");
  const selectedTemplate = templateSelect.value;
  let options;

  switch (selectedTemplate) {
    case "template1":
      options = template1Options;
      break;
    case "template2":
      options = template2Options;
      break;
    case "template3":
      options = template3Options;
      break;
    default:
      clearOptions();
      return;
  }

  updateOption1(options.option1);
  updateOption2(options.option2);
  updateOption3(options.option3);
  updateOption4(options.option4);
  updateOption5(options.option5);
}

function updateOption1(option) {
  const option1Container = document.getElementById("option1");
  option1Container.innerHTML = "";

  const checkbox = document.createElement("input");
  checkbox.type = "checkbox";
  checkbox.checked = option.checked;
  checkbox.id = "option1-checkbox";

  const label = document.createElement("label");
  label.textContent = option.text;
  label.htmlFor = "option1-checkbox";

  option1Container.appendChild(checkbox);
  option1Container.appendChild(label);
}

function updateOption2(option) {
  //https://stackoverflow.com/questions/17001961/how-to-add-drop-down-list-select-programmatically
  const option2Container = document.getElementById("option2");
  option2Container.innerHTML = "";
  const selectbox = document.createElement("select");
  option2Container.appendChild(selectbox);

  const optionElements = ["Cat", "Dog", "Wolf"];
  optionElements.forEach((optionElement) => {
    const option_add = document.createElement("option");
    option_add.value = optionElement;
    option_add.textContent = optionElement;
    option_add.selected = optionElement === option.value;
    selectbox.appendChild(option_add);
  });

  return;
  const option2Select = document.getElementById("option2");
  option2Select.innerHTML = "";

  const optionElements2 = ["Cat", "Dog", "Wolf"];
  optionElements.forEach((optionElement) => {
    const option_add = document.createElement("option");
    option_add.value = optionElement;
    option_add.textContent = optionElement;
    option_add.selected = optionElement === option.value;
    option2Select.appendChild(option_add);
  });
}

function updateOption3(option) {
  const option3Select = document.getElementById("option3");
  option3Select.innerHTML = "";

  const optionstr = option.values.toString();

  const optionElements = ["Lion", "Bird", "Butterfly"];
  optionElements.forEach((optionElement) => {
    const option = document.createElement("option");
    option.value = optionElement;
    option.textContent = optionElement;
    option.selected = optionstr.includes(optionElement);
    option3Select.appendChild(option);
  });
}

function updateOption4(option) {
  const option4Container = document.getElementById("option4");
  option4Container.innerHTML = "";

  const optionElements = ["Churu", "Cat Food", "Dog Food"];
  optionElements.forEach((optionElement) => {
    const radio = document.createElement("input");
    radio.type = "radio";
    radio.name = "option4";
    radio.value = optionElement;
    radio.id = `option4-${optionElement.replace(/\s/g, "-")}`;
    radio.checked = optionElement === option.value;

    const label = document.createElement("label");
    label.textContent = optionElement;
    label.htmlFor = radio.id;

    option4Container.appendChild(radio);
    option4Container.appendChild(label);
    option4Container.appendChild(document.createElement("br"));
  });
}

function updateOption5(option) {
  const option5Input = document.getElementById("option5");
  option5Input.value = option.text;
}

function clearOptions() {
  const optionContainers = [
    document.getElementById("option1"),
    document.getElementById("option2"),
    document.getElementById("option3"),
    document.getElementById("option4"),
  ];

  optionContainers.forEach((container) => {
    container.innerHTML = "";
  });

  const option5Input = document.getElementById("option5");
  option5Input.value = "";
}
