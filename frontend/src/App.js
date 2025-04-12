import React, { useRef, useState } from 'react';
import { Editor } from '@tinymce/tinymce-react';
import axios from 'axios';
import { saveAs } from 'file-saver';

function App() {
  const editorRef = useRef(null);
  const [loading, setLoading] = useState(false);

  const handleConvert = async () => {
    const html = editorRef.current.getContent();
    setLoading(true);

    try {
      const res = await axios.post('http://localhost:3001/convert', { html }, {
        responseType: 'blob',
      });

      saveAs(
        new Blob([res.data], {
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
        apiKey="1800qht7mdc06i65pnblmat83560kgdbma2n9t6a189pftbk"
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
    </div>
  );
}

export default App;
