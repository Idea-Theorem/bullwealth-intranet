import * as React from 'react';
import { useState, useRef } from 'react';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/folders';
import '@pnp/sp/files';

interface IImageUploadFieldProps {
  value: string;
  onChange: (value: string) => void;
  context: any;
}

export const ImageUploadField: React.FC<IImageUploadFieldProps> = ({ value, onChange, context }) => {
  const [uploading, setUploading] = useState(false);
  const [message, setMessage] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Initialize PnP correctly
  const sp = spfi().using(SPFx(context));

  const handleUpload = async (file: File): Promise<void> => {
    setUploading(true);
    setMessage('Uploading image...');

    try {
      const timestamp = new Date().getTime();
      const fileName = `news_${timestamp}_${file.name}`;
      const folderUrl = `/SiteAssets/NewsImages`;
      
      // Ensure folder exists
      try {
        await sp.web.getFolderByServerRelativePath(folderUrl)();
      } catch {
        await sp.web.folders.addUsingPath('SiteAssets/NewsImages');
      }
      
      // Upload file
      const folder = sp.web.getFolderByServerRelativePath(folderUrl);
      await folder.files.addUsingPath(fileName, file, { Overwrite: true });
      
      // Return full URL
      const siteUrl = context.pageContext.web.absoluteUrl;
      const uploadedUrl = `${siteUrl}/SiteAssets/NewsImages/${fileName}`;
      
      onChange(uploadedUrl);
      setMessage('Image uploaded successfully!');
      
    } catch (error) {
      console.error('Upload error:', error);
      setMessage('Upload failed. Please try again.');
    } finally {
      setUploading(false);
      setTimeout(() => setMessage(''), 3000);
    }
  };

  const handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (!file.type.startsWith('image/')) {
      setMessage('Please select an image file');
      setTimeout(() => setMessage(''), 3000);
      return;
    }

    if (file.size > 5 * 1024 * 1024) {
      setMessage('Image size should be less than 5MB');
      setTimeout(() => setMessage(''), 3000);
      return;
    }

    void handleUpload(file);
  };

  return (
    <div style={{ margin: '10px 0' }}>
      {message && (
        <MessageBar
          messageBarType={message.includes('success') ? MessageBarType.success : MessageBarType.error}
          isMultiline={false}
        >
          {message}
        </MessageBar>
      )}
      
      {value && (
        <div style={{ marginBottom: '10px' }}>
          <img
            src={value}
            alt="Preview"
            style={{ width: '120px', height: '80px', objectFit: 'cover', border: '1px solid #ccc', borderRadius: '4px' }}
          />
        </div>
      )}

      <div style={{ marginBottom: '10px' }}>
        <PrimaryButton
          text="📁 Upload Image"
          onClick={() => fileInputRef.current?.click()}
          disabled={uploading}
        />
        {value && (
          <DefaultButton
            text="Clear"
            onClick={() => onChange('')}
            style={{ marginLeft: '10px' }}
          />
        )}
      </div>

      <input
        type="file"
        accept="image/*"
        ref={fileInputRef}
        style={{ display: 'none' }}
        onChange={handleFileSelect}
      />

      <input
        type="text"
        placeholder="Or enter image URL"
        value={value}
        onChange={(e) => onChange(e.target.value)}
        style={{ width: '100%', padding: '8px', border: '1px solid #ccc', borderRadius: '4px' }}
      />
    </div>
  );
};
