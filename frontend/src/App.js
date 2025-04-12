
import React, { useRef, useState } from 'react';
import { Editor } from '@tinymce/tinymce-react';
import axios from 'axios';
import { saveAs } from 'file-saver';

function App() {
  const editorRef = useRef(null);
  const [loading, setLoading] = useState(false);
  const [ooxml, setOoxml] = useState('');

  const handleConvert = async () => {
    const html = editorRef.current.getContent();
    setLoading(true);
  
    try {
      // 1. Get the OOXML preview
      const xmlRes = await axios.post('https://html-to-ooxml.onrender.com/convert-preview', { html });
      if (xmlRes.data && xmlRes.data.ooxml) {
        setOoxml(xmlRes.data.ooxml);
      }
  
      // 2. Download actual .docx
      const docxRes = await axios.post('https://html-to-ooxml.onrender.com/convert-docx', { html }, {
        responseType: 'blob',
      });
  
      saveAs(
        new Blob([docxRes.data], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        }),
        'output.docx'
      );
    } catch (err) {
      alert('Conversion failed!');
      console.error(err);
    }
  
    setLoading(false);
  };
  return (
    <div style={{ padding: 20 }}>
      <h2>HTML to .docx Converter (Styled)</h2>

      <Editor
        apiKey="pfslezhfceo05a2eada79zz1b5fzlq1bqrcamv0r3b3wqs4i"
        onInit={(evt, editor) => (editorRef.current = editor)}
        initialValue={`<p>This is an <strong>example</strong> paragraph. <span style="color:red; font-size:18px;"><strong>Red bold text</strong></span></p>`}
        init={{
          height: 500,
          menubar: true,
          plugins: 'lists link code textcolor',
          toolbar:
            'undo redo | formatselect | fontselect fontsizeselect | bold italic underline | forecolor backcolor | alignleft aligncenter alignright | bullist numlist | link | code',
          fontsize_formats: '8px 10px 12px 14px 16px 18px 24px 36px',
          content_style: 'body { font-family:Arial,sans-serif; font-size:14px }',
        }}
      />

      <button
        onClick={handleConvert}
        disabled={loading}
        style={{
          marginTop: 20,
          padding: '10px 20px',
          fontSize: 16,
          background: '#007bff',
          color: '#fff',
          border: 'none',
          borderRadius: 5,
          cursor: 'pointer',
        }}
      >
        {loading ? 'Converting...' : 'Convert & Download'}
      </button>

      {ooxml && (
        <div style={{ marginTop: 30 }}>
          <h3>🔍 OOXML Preview</h3>
          <pre style={{
            background: '#f4f4f4',
            padding: 20,
            borderRadius: 5,
            overflowX: 'auto',
            maxHeight: 300,
            whiteSpace: 'pre-wrap'
          }}>
            {ooxml}
          </pre>
        </div>
      )}
    </div>
  );
}

export default App;
