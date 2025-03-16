'use client';

import { useState } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { Tabs, TabsList, TabsTrigger, TabsContent } from '@/components/ui/tabs';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Textarea } from '@/components/ui/textarea';
import { useDropzone } from 'react-dropzone';
import { TypeAnimation } from 'react-type-animation';
import { Upload, FileText, CheckCircle, AlertCircle } from 'lucide-react';

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
        setError('Failed to process file');
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
      setError('Failed to generate document');
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
    <div className="min-h-screen bg-white">
      {/* Header */}
      <header className="border-b">
        <div className="container mx-auto px-4 h-16 flex items-center">
          <motion.div
            initial={{ opacity: 0, x: -20 }}
            animate={{ opacity: 1, x: 0 }}
            className="flex items-center space-x-2"
          >
            <div className="font-bold text-xl text-blue-600">Cybergen-Doc</div>
          </motion.div>
        </div>
      </header>

      <main className="container mx-auto px-4 py-8">
        {/* About Section with Text Animation */}
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="mb-12 max-w-2xl mx-auto text-center"
        >
          <TypeAnimation
            sequence={[
              'Transform your documents with professional formatting',
              1000,
              'Automatic heading detection and smart date formatting',
              1000,
              'Preserve tables and images with consistent styling',
              1000,
            ]}
            wrapper="h2"
            speed={50}
            className="text-2xl font-semibold text-gray-800 mb-4"
            repeat={Infinity}
          />
        </motion.div>

        {/* Main Content */}
        <div className="max-w-4xl mx-auto">
          <Tabs defaultValue="import" className="w-full">
            <TabsList className="grid w-full grid-cols-2 mb-8">
              <TabsTrigger value="import" className="flex items-center space-x-2">
                <Upload className="w-4 h-4" />
                <span>Import Document</span>
              </TabsTrigger>
              <TabsTrigger value="text" className="flex items-center space-x-2">
                <FileText className="w-4 h-4" />
                <span>Enter Text</span>
              </TabsTrigger>
            </TabsList>

            <TabsContent value="import">
              <div className="space-y-6">
                <div
                  {...getRootProps()}
                  className={`border-2 border-dashed rounded-lg p-8 text-center transition-all duration-200
                    ${isDragActive 
                      ? 'border-blue-500 bg-blue-50'
                      : 'border-gray-300 hover:border-blue-500 hover:bg-gray-50'
                    }`}
                >
                  <input {...getInputProps()} />
                  <Upload className="w-12 h-12 mx-auto text-gray-400 mb-4" />
                  <p className="text-gray-600">
                    {isDragActive
                      ? "Drop your file here..."
                      : "Drag and drop a file here, or click to select"}
                  </p>
                  <p className="text-sm text-gray-500 mt-2">
                    Supported formats: .docx, .pdf
                  </p>
                </div>
                <Input
                  placeholder="Output filename (optional)"
                  value={filename}
                  onChange={(e) => setFilename(e.target.value)}
                  className="max-w-sm"
                />
                <Button
                  onClick={() => setLoading(true)}
                  disabled={loading}
                  className="w-full"
                >
                  {loading ? 'Processing...' : 'Generate Document'}
                </Button>
              </div>
            </TabsContent>

            <TabsContent value="text">
              <form onSubmit={handleTextSubmit} className="space-y-6">
                <Textarea
                  value={text}
                  onChange={(e) => setText(e.target.value)}
                  placeholder="Enter your document text here..."
                  className="min-h-[300px]"
                />
                <Input
                  placeholder="Output filename (optional)"
                  value={filename}
                  onChange={(e) => setFilename(e.target.value)}
                  className="max-w-sm"
                />
                <Button
                  type="submit"
                  disabled={loading}
                  className="w-full"
                >
                  {loading ? 'Processing...' : 'Generate Document'}
                </Button>
              </form>
            </TabsContent>
          </Tabs>

          {/* Status Messages */}
          <AnimatePresence>
            {(error || success) && (
              <motion.div
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: 20 }}
                className={`mt-4 p-4 rounded-lg flex items-center space-x-2
                  ${error ? 'bg-red-50 text-red-700' : 'bg-green-50 text-green-700'}`}
              >
                {success ? (
                  <CheckCircle className="w-5 h-5 text-green-500" />
                ) : (
                  <AlertCircle className="w-5 h-5 text-red-500" />
                )}
                <span>{error || 'Document generated successfully!'}</span>
              </motion.div>
            )}
          </AnimatePresence>

          {/* Features Grid */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mt-12">
            {[
              {
                title: 'Smart Formatting',
                description: 'Automatic detection of headings, subheadings, and paragraphs',
              },
              {
                title: 'Date Handling',
                description: 'Intelligent date detection and standardization',
              },
              {
                title: 'Content Preservation',
                description: 'Maintains tables, images, and complex formatting',
              },
            ].map((feature, index) => (
              <motion.div
                key={feature.title}
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ delay: index * 0.1 }}
                className="bg-blue-50 p-6 rounded-lg"
              >
                <h3 className="font-semibold text-blue-900 mb-2">{feature.title}</h3>
                <p className="text-blue-700 text-sm">{feature.description}</p>
              </motion.div>
            ))}
          </div>
        </div>
      </main>
    </div>
  );
}
