import { useCallback, useEffect, useRef, useState } from "react";
import Quill from "quill";
import "quill/dist/quill.snow.css";
import { io } from "socket.io-client";
import { useParams } from "react-router-dom";
import { saveAs } from "file-saver";
import * as htmlDocx from "html-docx-js/dist/html-docx";
import * as ReactDOM from "react-dom";
import * as docx2html from "docx2html";

const SAVE_INTERVAL_MS = 2000;
const TOOLBAR_OPTIONS = [
  [{ header: [1, 2, 3, 4, 5, 6, false] }],
  [{ font: [] }],
  [{ list: "ordered" }, { list: "bullet" }],
  ["bold", "italic", "underline"],
  [{ color: [] }, { background: [] }],
  [{ script: "sub" }, { script: "super" }],
  [{ align: [] }],
  ["image", "blockquote", "code-block"],
  ["clean"],
];

/**
 * @param {HTMLDivElement} wrapper
 */
function saveDoc(wrapper) {
  if (!wrapper) return;

  const qlEditorList = wrapper.getElementsByClassName("ql-editor");

  console.log({ qlEditorList });

  const jEditor = qlEditorList[0];

  if (!jEditor) return;

  const html = jEditor.innerHTML;

  let finalHtml = '<html><head><meta charset="UTF-8"></head><body>';
  finalHtml += html;
  finalHtml += "</body></html>";

  const converted = htmlDocx.asBlob(finalHtml);
  const docName = "document.docx";

  console.log(converted);

  saveAs(converted, docName);
}

export default function TextEditor() {
  const { id: documentId } = useParams();
  const [socket, setSocket] = useState();
  const [quill, setQuill] = useState();

  const saveTimerRef = useRef();

  useEffect(() => {
    const s = io("http://localhost:3001");
    setSocket(s);

    return () => {
      s.disconnect();
    };
  }, []);

  useEffect(() => {
    if (socket == null || quill == null) return;

    socket.once("load-document", (document) => {
      quill.setContents(document);
      quill.enable();
    });

    socket.emit("get-document", documentId);
  }, [socket, quill, documentId]);

  useEffect(() => {
    if (socket == null || quill == null) return;

    const handler = (delta) => {
      quill.updateContents(delta);
    };
    socket.on("receive-changes", handler);

    return () => {
      socket.off("receive-changes", handler);
    };
  }, [socket, quill]);

  useEffect(() => {
    if (socket == null || quill == null) return;

    const handler = (delta, oldDelta, source) => {
      if (source !== "user") return;
      socket.emit("send-changes", delta);

      if (saveTimerRef.current) clearTimeout(saveTimerRef.current);

      saveTimerRef.current = setTimeout(() => {
        socket.emit("save-document", quill.getContents());

        clearTimeout(saveTimerRef.current);
        saveTimerRef.current = null;
      }, SAVE_INTERVAL_MS);
    };

    quill.on("text-change", handler);

    return () => {
      quill.off("text-change", handler);
    };
  }, [socket, quill]);

  /**
   * @type {ReturnType<typeof useRef<HTMLInputElement>>}
   */
  const fileInputRef = useRef();

  /**
   * @type {ReturnType<typeof useRef<HTMLFormElement>>}
   */
  const fileInputFormRef = useRef();

  const importDoc = () => {
    if (!fileInputRef.current) return;

    fileInputRef.current.click();
  };

  /**
   * @type {React.ComponentProps<'input'>['onChange']}
   */
  const handleNewImport = async (e) => {
    e.preventDefault();

    const files = e.target?.files;
    if (!files) return;

    const newFile = files[0];
    if (!newFile) return;

    console.log(newFile);

    const htmlObj = await docx2html(newFile);

    console.log({htmlObj});

    fileInputFormRef.current?.reset();
  };

  const wrapperRef = useCallback((wrapper) => {
    if (wrapper == null) return;

    wrapper.innerHTML = "";
    const editor = document.createElement("div");
    wrapper.append(editor);

    const q = new Quill(editor, {
      theme: "snow",
      modules: { toolbar: TOOLBAR_OPTIONS },
    });

    q.disable();
    q.setText("Loading...");

    setQuill(q);

    // append custom buttons to toolbar
    /**
     * @type {HTMLDivElement}
     */
    const toolbarEl = wrapper.getElementsByClassName("ql-toolbar")[0];

    if (!toolbarEl) return;

    const containers = [
      [
        document.createElement("span"),
        <button key="save" onClick={() => saveDoc(wrapper)}>
          Save
        </button>,
      ],
      [
        document.createElement("span"),
        <button key="import" onClick={() => importDoc()}>
          Import
        </button>,
      ],
    ];

    for (const [container, jsx] of containers) {
      container.classList.add("ql-formats");
      ReactDOM.render(jsx, container);
      toolbarEl.append(container);
    }
  }, []);

  return (
    <>
      <div>
        <div className="container" ref={wrapperRef}></div>
      </div>

      <form ref={fileInputFormRef} hidden>
        <input
          hidden
          type="file"
          accept=".doc,.docx,.xml,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          ref={fileInputRef}
          onChange={handleNewImport}
        />
      </form>
    </>
  );
}
