/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/no-use-before-define */
import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './EmployeeDirectory.module.scss';
import { IEmployeeDirectoryProps, IEmployee } from './IEmployeeDirectoryProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner } from '@fluentui/react/lib/Spinner';


const EmployeeDirectory: React.FC<IEmployeeDirectoryProps> = (props) => {
  const [sectionData, setSectionData] = useState<{ [key: string]: IEmployee[] }>({});
  const [currentPages, setCurrentPages] = useState<{ [key: string]: number }>({});
  const [loading, setLoading] = useState<boolean>(true);

  const cardsPerPage = props.maxEmployeesToShow || 5;

  useEffect(() => {
    const fetchEmployeesForSections = async () => {
      setLoading(true);
      const newSectionData: { [key: string]: IEmployee[] } = {};
      const newPaginationState: { [key: string]: number } = {};

      for (const section of props.sections) {
        if (section && section.listName && section.title) {
          try {
            // ✅ FIXED: Updated to match your actual column names
            const url = `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(section.listName)}')/items?$select=Id,Title,Email,CompanyName,GroupName&$top=50`;
            console.log(`Fetching data for ${section.title} from:`, url);
            
            const response: SPHttpClientResponse = await props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
            
            console.log(`Response status for ${section.title}:`, response.status);
            
            if (response.ok) {
              const data = await response.json();
              console.log(`Data received for ${section.title}:`, data);
              
              // ✅ FIXED: Map CompanyName to job title
              newSectionData[section.title] = Array.isArray(data.value) ? data.value.map((item: any, idx: number) => ({
                id: item.Id || idx + 1,
                name: item.Title || "Unknown Employee",
                title: item.CompanyName || item.GroupName || "Employee", // Use CompanyName or GroupName as job title
                email: item.Email || "employee@company.com",
                phone: "",
                profileImage: ''
              })) : [];
              
              console.log(`Processed ${newSectionData[section.title].length} employees for ${section.title}`);
            } else {
              const errorText = await response.text();
              console.error(`Error fetching ${section.title}:`, response.status, errorText);
              newSectionData[section.title] = [];
            }
            newPaginationState[section.title] = 0;
          } catch (error) {
            console.error(`Exception fetching ${section.title}:`, error);
            newSectionData[section.title] = [];
            newPaginationState[section.title] = 0;
          }
        }
      }

      setSectionData(newSectionData);
      setCurrentPages(newPaginationState);
      setLoading(false);
    };

    if (props.sections && props.sections.length > 0) {
      fetchEmployeesForSections();
    } else {
      setLoading(false);
    }
  }, [props.sections, props.context]);

  const getInitials = (name: string): string => {
    const nameParts = name.trim().split(' ');
    if (nameParts.length >= 2) {
      return (nameParts[0][0] + nameParts[nameParts.length - 1][0]).toUpperCase();
    } else if (nameParts.length === 1) {
      return nameParts[0].substring(0, 2).toUpperCase();
    }
    return 'NA';
  };

  const getAvatarColor = (name: string): string => {
    const colors = [
      '#0078D4', '#107C10', '#D83B01', '#8764B8', '#008272',
      '#CA5010', '#00BCF2', '#498205', '#C239B3', '#0063B1'
    ];
    const charCode = name.split('').reduce((acc, char) => acc + char.charCodeAt(0), 0);
    return colors[charCode % colors.length];
  };

  const handleContactClick = (email: string) => {
    window.open(`mailto:${email}`, "_blank");
  };

  const handlePageChange = (sectionTitle: string, pageNum: number) => {
    setCurrentPages(prev => ({ ...prev, [sectionTitle]: pageNum }));
  };

  const renderEmployeeCard = (employee: IEmployee) => {
    const initials = getInitials(employee.name);
    const avatarColor = getAvatarColor(employee.name);

    return (
      <div className={styles.employeeCard} key={employee.id}>
        <div className={styles.profileSection}>
          <div 
            className={styles.initialsAvatar}
            style={{ backgroundColor: avatarColor }}
          >
            {initials}
          </div>

          <div className={styles.employeeInfo}>
            <h3 className={styles.employeeName}>{employee.name}</h3>
          </div>
        </div>
        <p className={styles.employeeTitle}>{employee.title}</p>
        <div className={styles.contactInfo}>
          <div className={styles.contactItem}>
            <Icon iconName="Mail" className={styles.contactIcon} />
            <span className={styles.contactText}>{employee.email}</span>
          </div>
        </div>
        <div className={styles.actionSection}>
          <button 
            className={styles.contactButton} 
            onClick={() => handleContactClick(employee.email)} 
            type="button"
          >
            <Icon iconName="Mail" className={styles.buttonIcon} />
            Contact
          </button>
        </div>
      </div>
    );
  };

  const renderSection = (title: string, employees: IEmployee[]) => {
    const currentPage = currentPages[title] || 0;
    const pageCount = Math.ceil((employees?.length || 0) / cardsPerPage);
    const startIdx = currentPage * cardsPerPage;
    const currentEmployees = employees?.slice(startIdx, startIdx + cardsPerPage) || [];

    return (
      <div key={title} className={styles.departmentSection}>
        {props.sections.length > 1 && (
          <h2 className={styles.sectionTitle}>{title}</h2>
        )}
        
        {currentEmployees.length > 0 ? (
          <>
            <div className={styles.employeeGrid}>
              {currentEmployees.map(renderEmployeeCard)}
            </div>
            {pageCount > 1 && (
              <div className={styles.pagination}>
                {Array.from({ length: pageCount }).map((_, idx) => (
                  <button
                    key={idx}
                    className={`${styles.paginationDot} ${idx === currentPage ? styles.active : ''}`}
                    onClick={() => handlePageChange(title, idx)}
                    type="button"
                    aria-label={`Page ${idx + 1}`}
                  >
                    ●
                  </button>
                ))}
              </div>
            )}
          </>
        ) : (
          <div className={styles.noEmployees}>
            No employees found in {title}.
          </div>
        )}
      </div>
    );
  };

  if (loading) {
    return (
      <div className={styles.loaderContainer}>
        <Spinner label="Loading employees..." />
      </div>
    );
  }

  return (
    <div className={styles.employeeDirectory}>
      <div className={styles.header}>
        <h1>{props.title}</h1>
        {props.orgChartLink && (
          <a 
            href={props.orgChartLink} 
            className={styles.orgChartButton} 
            target="_blank" 
            rel="noopener noreferrer"
          >
            <Icon iconName="Org" />
            Org Chart
          </a>
        )}
      </div>
      {props.sections && props.sections.length > 0 ? (
        props.sections.map(section =>
          renderSection(section.title, sectionData[section.title] || [])
        )
      ) : (
        <div className={styles.noSections}>
          No sections configured. Please configure sections in the web part properties.
        </div>
      )}
    </div>
  );
};

export default EmployeeDirectory;
