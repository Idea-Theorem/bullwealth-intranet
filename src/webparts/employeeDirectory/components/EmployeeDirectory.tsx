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
            const url = `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(section.listName)}')/items?$select=Id,Title,JobTitle,Email,WorkPhone,PhotoUrl&$top=50`;
            const response: SPHttpClientResponse = await props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
            
            if (response.ok) {
              const data = await response.json();
              newSectionData[section.title] = Array.isArray(data.value) ? data.value.map((item: any, idx: number) => ({
                id: item.Id || idx + 1,
                name: item.Title || "Unknown Employee",
                title: item.JobTitle || "Senior Tax Accountant",
                email: item.Email || "employee@company.com",
                phone: item.WorkPhone || "555-123-4567",
                profileImage: item.PhotoUrl || ''
              })) : [];
            } else {
              newSectionData[section.title] = generateDemoEmployees(section.title, 8);
            }
            newPaginationState[section.title] = 0;
          } catch (error) {
            console.error(`Error fetching ${section.title}:`, error);
            newSectionData[section.title] = generateDemoEmployees(section.title, 8);
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

  const generateDemoEmployees = (department: string, count: number): IEmployee[] => {
    const names = [
      'John Doe', 'Ana Martinez', 'Brandon Hubah', 'Denisa Farrow', 'Michael Chen',
      'Sarah Wilson', 'David Brown', 'Emily Davis', 'Robert Taylor', 'Lisa Martinez'
    ];

    const jobTitles = department === 'Compliance' 
      ? ['Senior Tax Accountant', 'Compliance Officer', 'Audit Manager', 'Risk Analyst']
      : ['Senior Tax Accountant', 'Investment Analyst', 'Research Manager', 'Portfolio Manager'];

    return Array.from({ length: count }, (_, index) => ({
      id: index + 1,
      name: names[index % names.length],
      title: jobTitles[index % jobTitles.length],
      email: `${names[index % names.length].toLowerCase().replace(' ', '.')}@bullwealth.com`,
      phone: `555-123-45${60 + index}`,
      profileImage: ''
    }));
  };

  // ✅ NEW: Function to get initials from name
  const getInitials = (name: string): string => {
    const nameParts = name.trim().split(' ');
    if (nameParts.length >= 2) {
      return (nameParts[0][0] + nameParts[nameParts.length - 1][0]).toUpperCase();
    } else if (nameParts.length === 1) {
      return nameParts[0].substring(0, 2).toUpperCase();
    }
    return 'NA';
  };

  // ✅ NEW: Function to generate color based on name
  const getAvatarColor = (name: string): string => {
    const colors = [
      '#0078D4', // Blue
      '#107C10', // Green
      '#D83B01', // Orange
      '#8764B8', // Purple
      '#008272', // Teal
      '#CA5010', // Dark Orange
      '#00BCF2', // Light Blue
      '#498205', // Olive Green
      '#C239B3', // Magenta
      '#0063B1'  // Dark Blue
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
          {/* ✅ CHANGED: Show initials instead of image */}
          {employee.profileImage ? (
            <img
              className={styles.profileImage}
              src={employee.profileImage}
              alt={`${employee.name} profile`}
              onError={(e) => {
                // If image fails to load, hide the image and show initials
                const target = e.currentTarget;
                target.style.display = 'none';
              }}
            />
          ) : null}
          
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
    if (!employees || employees.length === 0) return null;

    const currentPage = currentPages[title] || 0;
    const pageCount = Math.ceil(employees.length / cardsPerPage);
    const startIdx = currentPage * cardsPerPage;
    const currentEmployees = employees.slice(startIdx, startIdx + cardsPerPage);

    return (
      <div key={title} className={styles.departmentSection}>
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
