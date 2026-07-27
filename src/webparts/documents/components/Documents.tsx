/* eslint-disable no-prototype-builtins */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/no-unused-vars */

import * as React from 'react';
import styles from './Documents.module.scss';
import { IDocumentsProps, IDocumentCategory } from './IDocumentsProps';
import { SPHttpClient } from '@microsoft/sp-http';
import { Spinner } from '@fluentui/react/lib/Spinner';

export interface IDocumentsState {
  categories: IDocumentCategory[];
  loading: boolean;
}

export default class Documents extends React.Component<IDocumentsProps, IDocumentsState> {
  constructor(props: IDocumentsProps) {
    super(props);
    this.state = {
      categories: props.categories,
      loading: props.isDynamicMode
    };
  }

  public componentDidMount(): void {
    if (this.props.isDynamicMode) {
      this.fetchFoldersFromLibrary();
    }
  }

  public componentDidUpdate(prevProps: IDocumentsProps): void {
    if (
      prevProps.isDynamicMode !== this.props.isDynamicMode ||
      prevProps.folderPath !== this.props.folderPath ||
      prevProps.documentLibraryName !== this.props.documentLibraryName
    ) {
      if (this.props.isDynamicMode) {
        this.fetchFoldersFromLibrary();
      }
    }
  }

  private fetchFoldersFromLibrary = async (): Promise<void> => {
  this.setState({ loading: true });

  try {
    const { context, documentLibraryName, folderPath } = this.props;
    const siteUrl = context.pageContext.web.absoluteUrl;
    const sitePath = context.pageContext.web.serverRelativeUrl;
    const cleanLibraryName = documentLibraryName.trim();
    const cleanFolderPath = folderPath.trim();
    const fullPath = `${sitePath}/${cleanLibraryName}/${cleanFolderPath}`.replace(/\/+/g, '/');

    console.log('📁 Fetching folders from:', fullPath);

    // ✅ Step 1: Get folders under the given path
    const foldersApiUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl(@path)/Folders?@path='${encodeURIComponent(
      fullPath
    )}'&$select=Name,ServerRelativeUrl`;

    const foldersResponse = await context.spHttpClient.get(
      foldersApiUrl,
      SPHttpClient.configurations.v1
    );

    if (!foldersResponse.ok) {
      console.error('❌ Folders API failed:', foldersResponse.status);
      this.setState({ categories: [], loading: false });
      return;
    }

    const foldersData = await foldersResponse.json();
    const folderList = foldersData.value || [];

    console.log('✅ Folders found:', folderList.length);
    folderList.forEach((f: any) => console.log(`  📁 ${f.Name}`));

    if (folderList.length === 0) {
      this.setState({ categories: [], loading: false });
      return;
    }

    // ✅ Step 2: Build OrderBy Map (fetch metadata for each folder)
    const orderByMap: { [key: string]: number } = {};

    for (const folder of folderList) {
      const folderName = folder.Name;
      const folderServerUrl = folder.ServerRelativeUrl;

      console.log(`📡 Checking OrderBy for folder: ${folderName}`);

      try {
        // ✅ FIX: Use ListItemAllFields for folder (no more filter failures)
        const folderItemUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(
          folderServerUrl
        )}')/ListItemAllFields?$select=OrderBy,Order,Order_x0020_By`;

        const folderItemResponse = await context.spHttpClient.get(
          folderItemUrl,
          SPHttpClient.configurations.v1
        );

        if (folderItemResponse.ok) {
          const folderItemData = await folderItemResponse.json();

          let orderValue: any = null;
          const possibleKeys = ['OrderBy', 'Order', 'Order_x0020_By'];

          for (const key of possibleKeys) {
            if (folderItemData[key] !== undefined && folderItemData[key] !== null) {
              orderValue = folderItemData[key];
              break;
            }
          }

          if (orderValue !== null) {
            const parsedOrder = parseInt(String(orderValue), 10);
            if (!isNaN(parsedOrder)) {
              orderByMap[folderName] = parsedOrder;
              console.log(`✅ ${folderName} → OrderBy = ${parsedOrder}`);
            }
          } else {
            console.warn(`⚠️ No OrderBy value found for ${folderName}`);
          }
        } else {
          console.warn(`⚠️ Could not fetch OrderBy for ${folderName}: ${folderItemResponse.status}`);
        }
      } catch (err) {
        console.warn(`⚠️ Error fetching OrderBy for ${folderName}`, err);
      }
    }

    console.log('📊 OrderBy Map:', orderByMap);

    // ✅ Step 3: Map folders to categories with OrderBy
    const categoriesWithOrder: IDocumentCategory[] = folderList.map((folder: any) => {
      const folderName = folder.Name;
      const orderBy = orderByMap[folderName] || 999;
      const libraryParam = `${cleanLibraryName}/${cleanFolderPath}/${folderName}`;

      return {
        id: folderName,
        title: folderName,
        folderName: folderName,
        orderBy: orderBy,
        imageData: '',
        libraryUrl: '',
        viewAllUrl: '',
        pageUrl: '',
        viewDocumentsText: 'View Documents',
        libraryParam: libraryParam
      };
    });

    // ✅ Step 4: Sort folders by OrderBy
    const sortedCategories = categoriesWithOrder.sort(
      (a, b) => (a.orderBy || 999) - (b.orderBy || 999)
    );

    const finalOrder = sortedCategories
      .map((c: IDocumentCategory) => `${c.title}(${c.orderBy})`)
      .join(' → ');

    console.log('✅ Final sorted folder order:', finalOrder);

    // ✅ Step 5: Update UI state
    this.setState({ categories: sortedCategories, loading: false });
  } catch (error) {
    console.error('❌ Error fetching folders:', error);
    this.setState({ categories: [], loading: false });
  }
};



  private handleViewDocuments = (category: IDocumentCategory): void => {
    try {
      if (category.libraryUrl && category.libraryUrl.trim() !== '') {
        window.open(category.libraryUrl, '_self');
        return;
      }

      if ((category as any).libraryParam) {
        const libraryParam = encodeURIComponent((category as any).libraryParam);
        const currentSiteUrl = this.props.context.pageContext.web.absoluteUrl;
        const docLibraryPageUrl = `${currentSiteUrl}/SitePages/DocumentLibrary.aspx`;
        const url = `${docLibraryPageUrl}?library=${libraryParam}`;
        window.location.href = url;
      }
    } catch (error) {
      console.error('Error navigating:', error);
    }
  };

  private getGradientForFolder = (index: number): { bg: string; icon: string } => {
    const colors = [
      { bg: '#2c3e50', icon: '#ffffff' },
      { bg: '#27ae60', icon: '#ffffff' },
      { bg: '#3498db', icon: '#ffffff' },
      { bg: '#2980b9', icon: '#ffffff' },
      { bg: '#e74c3c', icon: '#ffffff' },
      { bg: '#f39c12', icon: '#ffffff' }
    ];
    return colors[index % colors.length];
  };

  private getColorfulImage = (title: string, index: number): string => {
    const colors = this.getGradientForFolder(index);
    const encodedBg = colors.bg.replace('#', '%23');

    return `data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 400 280'%3E%3Crect width='400' height='280' fill='${encodedBg}'/%3E%3Cg%3E%3Crect x='130' y='80' width='140' height='140' rx='8' fill='none' stroke='white' stroke-width='3'/%3E%3Cline x1='150' y1='110' x2='250' y2='110' stroke='white' stroke-width='2.5' stroke-linecap='round'/%3E%3Cline x1='150' y1='140' x2='250' y2='140' stroke='white' stroke-width='2.5' stroke-linecap='round'/%3E%3Cline x1='150' y1='170' x2='210' y2='170' stroke='white' stroke-width='2.5' stroke-linecap='round'/%3E%3C/g%3E%3C/svg%3E`;
  };

  public render(): React.ReactElement {
    const { columnsPerRow } = this.props;
    const { categories, loading } = this.state;
    const gridClassName = (styles as any)[`columns${columnsPerRow}`] || styles.columns4;

    if (loading) {
      return <Spinner label="Loading folders..." />;
    }

    if (categories.length === 0) {
      return (
        <div style={{ padding: '20px', textAlign: 'center', color: '#666' }}>
          No folders found in the library.
        </div>
      );
    }

    return (
      <div className={`${styles.categoriesGrid} ${gridClassName}`}>
        {categories.map((category, index) => (
          <div
            key={category.id}
            style={{
              backgroundColor: '#ffffff',
              borderRadius: '12px',
              overflow: 'hidden',
              boxShadow: '0 2px 8px rgba(0, 0, 0, 0.1)',
              transition: 'all 0.3s ease',
              cursor: 'pointer'
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.transform = 'translateY(-4px)';
              e.currentTarget.style.boxShadow = '0 8px 16px rgba(0, 0, 0, 0.15)';
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.transform = 'translateY(0)';
              e.currentTarget.style.boxShadow = '0 2px 8px rgba(0, 0, 0, 0.1)';
            }}
          >
            <div
              style={{
                width: '100%',
                height: '200px',
                overflow: 'hidden',
                position: 'relative',
                backgroundColor: '#f5f5f5'
              }}
            >
              <img
                src={category.imageData || this.getColorfulImage(category.title, index)}
                alt={category.title}
                style={{
                  width: '100%',
                  height: '100%',
                  objectFit: 'cover',
                  transition: 'transform 0.3s ease'
                }}
                onMouseEnter={(e) => {
                  e.currentTarget.style.transform = 'scale(1.05)';
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.transform = 'scale(1)';
                }}
              />
            </div>

            <div style={{ padding: '20px' }}>
              <h3
                style={{
                  margin: '0 0 12px 0',
                  fontSize: '18px',
                  fontWeight: '700',
                  color: '#1a1a1a',
                  lineHeight: '1.3'
                }}
              >
                {category.title}
              </h3>

              <a
                onClick={(e) => {
                  e.preventDefault();
                  this.handleViewDocuments(category);
                }}
                style={{
                  display: 'inline-block',
                  color: '#2ecc71',
                  fontSize: '14px',
                  fontWeight: '600',
                  textDecoration: 'none',
                  cursor: 'pointer',
                  transition: 'all 0.2s ease'
                }}
                onMouseEnter={(e) => {
                  e.currentTarget.style.color = '#27ae60';
                  e.currentTarget.style.textDecoration = 'underline';
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.color = '#2ecc71';
                  e.currentTarget.style.textDecoration = 'none';
                }}
              >
                {category.viewDocumentsText || 'View Documents'}
              </a>
            </div>
          </div>
        ))}
      </div>
    );
  }
}
