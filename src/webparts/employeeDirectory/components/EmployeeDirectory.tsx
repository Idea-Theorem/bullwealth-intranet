import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './EmployeeDirectory.module.scss';
import { IEmployeeDirectoryProps, IEmployee } from './IEmployeeDirectoryProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner } from '@fluentui/react/lib/Spinner';

const EmployeeDirectory: React.FC<IEmployeeDirectoryProps> = (props) => {
  const [allEmployees, setAllEmployees] = useState<IEmployee[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [currentPage, setCurrentPage] = useState<number>(0);
  const employeesPerPage = 4; // Show 4 employees per page

  useEffect(() => {
    // eslint-disable-next-line @typescript-eslint/no-use-before-define
    fetchEmployees().catch((err) => {
      console.error('Error in fetchEmployees:', err);
    });
  }, []);

  const fetchEmployees = async (): Promise<void> => {
    try {
      setLoading(true);
      
      // Try to get from SharePoint User Information List
      const usersUrl = `${props.context.pageContext.web.absoluteUrl}/_api/web/siteusers?$filter=PrincipalType eq 1&$select=Id,Title,Email,JobTitle,Picture&$top=50`;
      
      const response: SPHttpClientResponse = await props.context.spHttpClient.get(
        usersUrl,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        
        if (data.value && data.value.length > 0) {
          // Map SharePoint users to our employee interface
          const sharePointEmployees: IEmployee[] = data.value
            .filter((user: any) => user.Email && user.Email.indexOf('@') > -1) // Only users with valid emails
            .map((user: any, index: number) => ({
              id: user.Id || index + 1,
              name: user.Title || 'Unknown User',
              title: user.JobTitle || 'Employee',
              email: user.Email,
              phone: '555-123-4567', // Default phone since SP doesn't have this in user info
              profileImage: user.Picture || `/_layouts/15/userphoto.aspx?size=L&username=${user.Email}`,
              department: 'General'
            }));

          setAllEmployees(sharePointEmployees);
        } else {
          // Fallback to demo data if no SharePoint users found
          throw new Error('No users found in SharePoint');
        }
      } else {
        throw new Error('Failed to fetch from SharePoint');
      }
    } catch (err: any) {
      console.error('Error fetching employees, using demo data:', err);
      
      // Fallback demo data - more realistic variety
      const demoEmployees: IEmployee[] = [
        {
          id: 1,
          name: 'Sarah Johnson',
          title: 'Senior Tax Accountant',
          email: 'sarah.johnson@company.com',
          phone: '555-123-4567',
          profileImage: `https://randomuser.me/api/portraits/women/44.jpg`,
          department: 'Finance'
        },
        {
          id: 2,
          name: 'Michael Chen',
          title: 'Financial Analyst',
          email: 'michael.chen@company.com',
          phone: '555-123-4568',
          profileImage: `https://randomuser.me/api/portraits/men/32.jpg`,
          department: 'Finance'
        },
        {
          id: 3,
          name: 'Emily Davis',
          title: 'Compliance Manager',
          email: 'emily.davis@company.com',
          phone: '555-123-4569',
          profileImage: `https://randomuser.me/api/portraits/women/68.jpg`,
          department: 'Compliance'
        },
        {
          id: 4,
          name: 'Robert Wilson',
          title: 'Senior Auditor',
          email: 'robert.wilson@company.com',
          phone: '555-123-4570',
          profileImage: `https://randomuser.me/api/portraits/men/75.jpg`,
          department: 'Audit'
        },
        {
          id: 5,
          name: 'Lisa Martinez',
          title: 'Risk Analyst',
          email: 'lisa.martinez@company.com',
          phone: '555-123-4571',
          profileImage: `https://randomuser.me/api/portraits/women/55.jpg`,
          department: 'Risk Management'
        },
        {
          id: 6,
          name: 'David Thompson',
          title: 'Tax Specialist',
          email: 'david.thompson@company.com',
          phone: '555-123-4572',
          profileImage: `https://randomuser.me/api/portraits/men/41.jpg`,
          department: 'Finance'
        },
        {
          id: 7,
          name: 'Jennifer Brown',
          title: 'Compliance Officer',
          email: 'jennifer.brown@company.com',
          phone: '555-123-4573',
          profileImage: `https://randomuser.me/api/portraits/women/22.jpg`,
          department: 'Compliance'
        },
        {
          id: 8,
          name: 'James Rodriguez',
          title: 'Financial Controller',
          email: 'james.rodriguez@company.com',
          phone: '555-123-4574',
          profileImage: `https://randomuser.me/api/portraits/men/18.jpg`,
          department: 'Finance'
        }
      ];

      setAllEmployees(demoEmployees);
    } finally {
      setLoading(false);
    }
  };

  const handleContact = (employee: IEmployee): void => {
    window.open(`mailto:${employee.email}`, '_blank');
  };

  const handlePageChange = (pageIndex: number): void => {
    setCurrentPage(pageIndex);
  };

  // Calculate pagination
  const totalPages = Math.ceil(allEmployees.length / employeesPerPage);
  const startIndex = currentPage * employeesPerPage;
  const currentEmployees = allEmployees.slice(startIndex, startIndex + employeesPerPage);

  if (loading) {
    return (
      <div className={styles.employeeDirectory}>
        <div className={styles.header}>
          <h2>{props.title}</h2>
        </div>
        <div className={styles.loading}>
          <Spinner label="Loading employees..." />
        </div>
      </div>
    );
  }

  return (
    <div className={styles.employeeCard} key={employee.id}>
      <div className={styles.profileSection}>
        <img
          className={styles.profileImage}
          src={employee.profileImage}
          alt={`${employee.name} profile`}
          onError={(e) => {
            const target = e.currentTarget;
            target.src = `https://ui-avatars.com/api/?name=${encodeURIComponent(employee.name)}&background=E1E1E1&color=666666`;
          }}
        />
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
        {/* <div className={styles.contactItem}>
          <Icon iconName="Phone" className={styles.contactIcon} />
          <span className={styles.contactText}>{employee.phone}</span>
        </div> */}
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
        <h2 className={styles.departmentTitle}>{title}</h2>
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
      
      {totalPages > 1 && (
        <div className={styles.pagination}>
          {Array.from({ length: totalPages }, (_, index) => (
            <button
              key={index}
              className={`${styles.paginationDot} ${index === currentPage ? styles.active : ''}`}
              onClick={() => handlePageChange(index)}
              title={`Page ${index + 1}`}
              aria-label={`Go to page ${index + 1}`}
            />
          ))}
        </div>
      )}
    </div>
  );
};

export default EmployeeDirectory;