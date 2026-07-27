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
  const [employees, setEmployees] = useState<IEmployee[]>([]);
  const [currentPage, setCurrentPage] = useState<number>(0);
  const [loading, setLoading] = useState<boolean>(true);

  const cardsPerPage = props.maxEmployeesToShow || 5;

  useEffect(() => {
    const fetchEmployees = async () => {
      setLoading(true);
      try {
        // ✅ Build filter query based on selected company
        let filterQuery = '';
        if (props.selectedCompany && props.selectedCompany !== 'All') {
          filterQuery = `&$filter=CompanyName eq '${props.selectedCompany.replace(/'/g, "''")}'`;
        }

        const url = `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(props.listName)}')/items?$select=Id,Title,Email,CompanyName,GroupName&$top=1000${filterQuery}`;
        console.log('Fetching employees from:', url);
        
        const response: SPHttpClientResponse = await props.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        
        if (response.ok) {
          const data = await response.json();
          console.log('Data received:', data);
          
          // ✅ Map and sort alphabetically
          const employeeList: IEmployee[] = Array.isArray(data.value) 
            ? data.value
                .map((item: any, idx: number) => ({
                  id: item.Id || idx + 1,
                  name: item.Title || "Unknown Employee",
                  title: item.CompanyName || "Employee",
                  email: item.Email || "employee@company.com",
                  phone: "",
                  profileImage: '',
                  groupName: item.GroupName || "General"
                }))
                .sort((a: IEmployee, b: IEmployee) => a.name.localeCompare(b.name))
            : [];
          
          console.log(`Processed ${employeeList.length} employees`);
          setEmployees(employeeList);
          
        } else {
          const errorText = await response.text();
          console.error('Error fetching employees:', response.status, errorText);
          setEmployees([]);
        }
      } catch (error) {
        console.error('Exception fetching employees:', error);
        setEmployees([]);
      }
      
      setLoading(false);
    };

    if (props.listName) {
      fetchEmployees();
    } else {
      setLoading(false);
    }
  }, [props.listName, props.context, props.selectedCompany]);

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

  const handlePageChange = (pageNum: number) => {
    setCurrentPage(pageNum);
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

  if (loading) {
    return (
      <div className={styles.loaderContainer}>
        <Spinner label="Loading employees..." />
      </div>
    );
  }

  // ✅ Calculate pagination
  const pageCount = Math.ceil(employees.length / cardsPerPage);
  const startIdx = currentPage * cardsPerPage;
  const currentEmployees = employees.slice(startIdx, startIdx + cardsPerPage);

  return (
    <div className={styles.employeeDirectory}>
      <div className={styles.header}>
        <h1>{props.title}</h1>
        {props.orgChartLink && (
          <a 
            href={props.orgChartLink} 
            className={styles.orgChartLink}
            target="_blank" 
            rel="noopener noreferrer"
          >
            View Organization Chart
          </a>
        )}
      </div>

      {/* ✅ REMOVED: Frontend filter dropdown */}

      {currentEmployees.length > 0 ? (
        <>
          <div className={styles.employeeGrid}>
            {currentEmployees.map(renderEmployeeCard)}
          </div>
          
          {pageCount > 1 && (
            <div className={styles.pagination}>
              {Array.from({ length: pageCount }).map((_, idx: number) => (
                <button
                  key={idx}
                  className={`${styles.paginationDot} ${idx === currentPage ? styles.active : ''}`}
                  onClick={() => handlePageChange(idx)}
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
          {props.selectedCompany && props.selectedCompany !== 'All' 
            ? `No employees found for ${props.selectedCompany}.` 
            : 'No employees found.'}
        </div>
      )}
    </div>
  );
};

export default EmployeeDirectory;
