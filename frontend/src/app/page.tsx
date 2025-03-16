'use client';

import { useState } from 'react';
import { Tab } from '@headlessui/react';
import { useDropzone } from 'react-dropzone';
import { motion } from 'framer-motion';
import { DocumentTextIcon, DocumentArrowUpIcon, CheckCircleIcon } from '@heroicons/react/24/outline';

export default function Home() {
  const [text, setText] = useState('');
  const [filename, setFilename] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState(false);

  const onDrop = async (acceptedFiles: File[]) => {
    if (acceptedFiles.length > 0) {
      const file = acceptedFiles[0];
      const formData = new FormData();
      formData.append('file', file);
      formData.append('filename', filename || 'generated_document.docx');
      
      try {
        setLoading(true);
        setError('');
        setSuccess(false);
        
        const response = await fetch('/api/generate-from-file', {
          method: 'POST',
          body: formData,
        });
        
        const data = await response.json();
        
        if (data.success) {
          setSuccess(true);
          downloadDocument(data.document, filename || 'generated_document.docx');
        } else {
          setError(data.error || 'Failed to process file');
        }
      } catch (err) {
        setError('Failed to process file' +err);
      } finally {
        setLoading(false);
      }
    }
  };

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'],
      'application/pdf': ['.pdf'],
    },
  });

  const handleTextSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    
    try {
      setLoading(true);
      setError('');
      setSuccess(false);
      
      const response = await fetch('/api/generate-from-text', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          text,
          filename: filename || 'generated_document.docx',
        }),
      });
      
      const data = await response.json();
      
      if (data.success) {
        setSuccess(true);
        downloadDocument(data.document, filename || 'generated_document.docx');
      } else {
        setError(data.error || 'Failed to generate document');
      }
    } catch (err) {
      setError('Failed to generate document' + err);
    } finally {
      setLoading(false);
    }
  };

  const downloadDocument = (base64Content: string, filename: string) => {
    const link = document.createElement('a');
    link.href = `data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,${base64Content}`;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div className="py-10">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="text-center">
          <motion.h1 
            className="text-4xl font-bold text-gray-900 mb-2"
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5 }}
          >
            Document Formatter
          </motion.h1>
          <motion.p 
            className="text-lg text-gray-600 mb-8"
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5, delay: 0.1 }}
          >
            Transform your documents with professional formatting
          </motion.p>
        </div>

        <motion.div
          className="bg-white rounded-2xl shadow-xl overflow-hidden"
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.5, delay: 0.2 }}
        >
          <Tab.Group>
            <Tab.List className="flex p-1 space-x-1 bg-blue-900/5">
              <Tab className={({ selected }) =>
                `w-full py-3 text-sm font-medium leading-5 rounded-lg transition-all duration-200
                ${selected 
                  ? 'bg-white text-blue-700 shadow'
                  : 'text-gray-600 hover:bg-white/[0.12] hover:text-blue-600'
                }`
              }>
                <div className="flex items-center justify-center space-x-2">
                  <DocumentTextIcon className="w-5 h-5" />
                  <span>Enter Text</span>
                </div>
              </Tab>
              <Tab className={({ selected }) =>
                `w-full py-3 text-sm font-medium leading-5 rounded-lg transition-all duration-200
                ${selected 
                  ? 'bg-white text-blue-700 shadow'
                  : 'text-gray-600 hover:bg-white/[0.12] hover:text-blue-600'
                }`
              }>
                <div className="flex items-center justify-center space-x-2">
                  <DocumentArrowUpIcon className="w-5 h-5" />
                  <span>Import Document</span>
                </div>
              </Tab>
            </Tab.List>
            <Tab.Panels className="p-6">
              <Tab.Panel>
                <form onSubmit={handleTextSubmit} className="space-y-4">
                  <div>
                    <label htmlFor="text" className="block text-sm font-medium text-gray-700 mb-1">
                      Document Text
                    </label>
                    <textarea
                      id="text"
                      value={text}
                      onChange={(e) => setText(e.target.value)}
                      className="w-full h-64 p-4 rounded-lg border border-gray-300 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                      placeholder="Enter your document text here..."
                    />
                  </div>
                  <div>
                    <label htmlFor="filename" className="block text-sm font-medium text-gray-700 mb-1">
                      Output Filename
                    </label>
                    <input
                      id="filename"
                      type="text"
                      value={filename}
                      onChange={(e) => setFilename(e.target.value)}
                      className="w-full p-2 rounded-lg border border-gray-300 focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                      placeholder="generated_document.docx"
                    />
                  </div>
                  <button
                    type="submit"
                    disabled={loading}
                    className={`w-full py-3 rounded-lg text-white font-medium transition-all duration-200
                      ${loading 
                        ? 'bg-blue-400 cursor-not-allowed'
                        : 'bg-blue-600 hover:bg-blue-700'
                      }`}
                  >
                    {loading ? 'Generating...' : 'Generate Document'}
                  </button>
                </form>
              </Tab.Panel>
              <Tab.Panel>
                <div
                  {...getRootProps()}
                  className={`border-2 border-dashed rounded-lg p-8 text-center transition-all duration-200
                    ${isDragActive 
                      ? 'border-blue-500 bg-blue-50'
                      : 'border-gray-300 hover:border-blue-500 hover:bg-gray-50'
                    }`}
                >
                  <input {...getInputProps()} />
                  <DocumentArrowUpIcon className="w-12 h-12 mx-auto text-gray-400 mb-4" />
                  <p className="text-gray-600">
                    {isDragActive
                      ? "Drop your file here..."
                      : "Drag and drop a file here, or click to select"}
                  </p>
                  <p className="text-sm text-gray-500 mt-2">
                    Supported formats: .docx, .pdf
                  </p>
                </div>
              </Tab.Panel>
            </Tab.Panels>
          </Tab.Group>
        </motion.div>

        {(error || success) && (
          <motion.div
            className={`mt-4 p-4 rounded-lg flex items-center space-x-2
              ${error ? 'bg-red-50 text-red-700' : 'bg-green-50 text-green-700'}`}
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.3 }}
          >
            {success ? (
              <CheckCircleIcon className="w-5 h-5 text-green-500" />
            ) : null}
            <span>{error || 'Document generated successfully!'}</span>
          </motion.div>
        )}

        <div className="mt-12">
          <h2 className="text-2xl font-bold text-gray-900 mb-4">Features</h2>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            {[
              {
                title: 'Smart Formatting',
                description: 'Automatic detection and formatting of headings, subheadings, and paragraphs',
              },
              {
                title: 'Date Handling',
                description: 'Intelligent date detection and standardization to a professional format',
              },
              {
                title: 'Content Preservation',
                description: 'Maintains tables, images, and complex formatting from source documents',
              },
            ].map((feature, index) => (
              <motion.div
                key={index}
                className="bg-white p-6 rounded-lg shadow-sm"
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ duration: 0.5, delay: 0.1 * index }}
              >
                <h3 className="text-lg font-semibold text-gray-900 mb-2">
                  {feature.title}
                </h3>
                <p className="text-gray-600">
                  {feature.description}
                </p>
              </motion.div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}
