import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './AdminDoc.module.scss';
import { IAdminDocProps } from './IAdminDocProps';
import { Nav, INavLink, INavStyles } from '@fluentui/react/lib/Nav';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { IconButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { Link } from '@fluentui/react/lib/Link';
import { ProgressIndicator } from '@fluentui/react/lib/ProgressIndicator';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { DocumentsView } from './DocumentsView';
import { CreateDocumentView } from './CreateDocumentView';

const navLinks = [
  { links: [
      { name: 'Dashboard', url: '', key: 'dashboard', icon: 'Home' },
      { name: 'Documents', url: '', key: 'documents', icon: 'Document' },
      { name: 'My Tasks', url: '', key: 'tasks', icon: 'TaskManager' },
      { name: 'Create Document', url: '', key: 'create', icon: 'Add' }
  ]},
  { links: [
      { name: 'Templates', url: '', key: 'templates', icon: 'PageTemplates' },
      { name: 'Workflow Configuration', url: '', key: 'workflow', icon: 'Flow' },
      { name: 'Reports', url: '', key: 'reports', icon: 'ReportDocument' }
  ]},
  { name: 'Administration', links: [
      { name: 'Document Types', url: '', key: 'doc-types', icon: 'DocumentSet' },
      { name: 'Metadata', url: '', key: 'metadata', icon: 'Tag' }, 
      { name: 'Numbering', url: '', key: 'numbering', icon: 'NumberSymbol' },
      { name: 'Users & Permissions', url: '', key: 'users', icon: 'People' },
      { name: 'Settings', url: '', key: 'settings', icon: 'Settings' }
  ]}
];

const navStyles: Partial<INavStyles> = { root: { height: '100%' }, groupContent: { marginBottom: '20px' } };

export default function AdminDoc(props: IAdminDocProps) {
  const [activeScreen, setActiveScreen] = useState<string>('dashboard');
  const [department, setDepartment] = useState<string>('Loading...');

  useEffect(() => {
    const fetchUserProfile = async () => {
      try {
        const response: SPHttpClientResponse = await props.context.spHttpClient.get(
          `${props.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`,
          SPHttpClient.configurations.v1
        );
        const data = await response.json();
        const deptProperty = data.UserProfileProperties.find((prop: any) => prop.Key === 'Department');
        const titleProperty = data.UserProfileProperties.find((prop: any) => prop.Key === 'Title');
        setDepartment(deptProperty?.Value || titleProperty?.Value || 'No Department');
      } catch (error) {
        setDepartment('Profile unavailable');
      }
    };
    fetchUserProfile();
  }, [props.context]);

  const handleNavClick = (ev?: React.MouseEvent<HTMLElement>, item?: INavLink) => {
    ev?.preventDefault(); 
    if (item && item.key) setActiveScreen(item.key);
  };

  // Helper to generate the 5-item lists
  const renderList = (title: string, items: any[]) => (
    <div className={styles.card}>
      <div className={styles.listHeader}>
        <span>{title}</span>
        <Link>View all</Link>
      </div>
      {items.map((item, i) => (
        <div key={i} className={styles.listItem}>
          <Icon iconName={item.icon} className={styles.listIcon} />
          <span className={styles.listName}>{item.name}</span>
          <span className={styles.listStatus}>{item.status}</span>
          <span className={styles.listDue}>{item.due}</span>
        </div>
      ))}
      <div className={styles.listFooter}>
        <span>Showing 1 to 5 of 24 tasks</span>
        <div>
          <IconButton iconProps={{ iconName: 'ChevronLeft' }} />
          <IconButton iconProps={{ iconName: 'ChevronRight' }} />
        </div>
      </div>
    </div>
  );

  const renderDashboard = () => {
    // Mock data for the lists
    const mockTasks = Array(5).fill({ name: 'Approve HR Policy', status: 'Pending', due: 'Today', icon: 'WordDocument' });
    const mockExpiring = Array(5).fill({ name: 'Vendor Contract NDA', status: 'Expiring', due: 'In 3 Days', icon: 'PDFDocument' });
    const mockRecent = Array(5).fill({ name: 'Q3 Financial Report', status: 'Draft', due: '-', icon: 'ExcelDocument' });

    return (
      <div>
        {/* Welcome Row */}
        <div className={styles.dashboardHeader}>
          <h2>Welcome, {props.userDisplayName}</h2>
          <PrimaryButton text="Add new document" iconProps={{ iconName: 'Add' }} />
        </div>

        {/* 5 Top Summary Cards */}
        <div className={styles.gridRow5}>
          <div className={`${styles.card} ${styles.summaryCard}`}><div className={styles.metric}>1,402</div><div className={styles.metricLabel}>Total Documents</div></div>
          <div className={`${styles.card} ${styles.summaryCard}`}><div className={styles.metric}>24</div><div className={styles.metricLabel}>Total My Tasks</div></div>
          <div className={`${styles.card} ${styles.summaryCard}`}><div className={styles.metric}>8</div><div className={styles.metricLabel}>Expiring Soon</div></div>
          <div className={`${styles.card} ${styles.summaryCard}`}><div className={styles.metric}>954</div><div className={styles.metricLabel}>Approved Documents</div></div>
          <div className={`${styles.card} ${styles.summaryCard}`}><div className={styles.metric}>18</div><div className={styles.metricLabel}>Pending Signature</div></div>
        </div>

        {/* 3 Middle Cards (Lists) */}
        <div className={styles.gridRow3}>
          {renderList("My Tasks", mockTasks)}
          {renderList("Expiring / Review Due", mockExpiring)}
          {renderList("Recent Documents", mockRecent)}
        </div>

        {/* 3 Bottom Cards */}
        <div className={styles.gridRow3}>
          
          {/* Donut Chart Card */}
          <div className={styles.card}>
            <div className={styles.listHeader}><span>Document Status</span></div>
            <div className={styles.donutWrapper}>
              <div className={styles.donut}></div>
              <div className={styles.donutLegend}>
                <div>🟢 Draft: 32 (26%)</div>
                <div>🟠 Under Review: 28 (23%)</div>
                <div>🔴 Pending Approval: 18 (15%)</div>
                <div>✅ Approve: 78 (31%)</div>
                <div>🔵 Published: 78 (31%)</div>
                <div>⚫ Archived: 12 (10%)</div>
              </div>
            </div>
          </div>

          {/* Progress Indicator Card */}
          <div className={styles.card}>
            <div className={styles.listHeader}><span>Document by Type</span></div>
            <div>
              <ProgressIndicator label="SOP" description="54 (44%)" percentComplete={0.44} />
              <ProgressIndicator label="Policy" description="28 (23%)" percentComplete={0.23} />
              <ProgressIndicator label="Contract" description="20 (16%)" percentComplete={0.16} />
              <ProgressIndicator label="Form" description="8 (5%)" percentComplete={0.05} />
              <ProgressIndicator label="Work Instruction" description="8 (5%)" percentComplete={0.05} />
              <ProgressIndicator label="Others" description="8 (5%)" percentComplete={0.05} />
            </div>
          </div>

          {/* Announcements Card */}
          <div className={styles.card}>
            <div className={styles.listHeader}><span>Announcements</span></div>
            <div className={styles.announcementItem}>
              <Icon iconName="MegaphoneSolid" className={styles.announceIcon} />
              <div>
                <div className={styles.announceTitle}>System Maintenance</div>
                <div className={styles.announceDesc}>SharePoint environment will undergo scheduled maintenance to upgrade database performance.</div>
                <div className={styles.announceDesc}><strong>Date:</strong> Aug 22, 10:00 PM - Aug 23, 02:00 AM</div>
                <div className={styles.announceMeta}>Published: Aug 17, 2026 - IT Infrastructure Dept</div>
              </div>
            </div>
          </div>

        </div>
      </div>
    );
  };

  const renderContent = () => {
  switch (activeScreen) {
    case 'dashboard': return renderDashboard();
    case 'documents': return <DocumentsView />;
    case 'create': return <CreateDocumentView context={props.context} userDisplayName={props.userDisplayName} onSuccess={() => setActiveScreen('documents')} />;
    default: return <div><h2>Under Construction</h2><p>This screen is not built yet.</p></div>;
  }
};

  return (
    <div className={styles.appWrapper}>
      <div className={styles.topHeader}>
        <div className={styles.headerLeft}><span>DocuPro</span></div>
        <div className={styles.headerRight}>
          <IconButton iconProps={{ iconName: 'Ringer' }} title="Notifications" />
          <IconButton iconProps={{ iconName: 'Help' }} title="Help" />
          <Persona text={props.userDisplayName} secondaryText={department} size={PersonaSize.size32} />
        </div>
      </div>
      <div className={styles.adminDoc}>
        <div className={styles.sidebar}>
          <Nav groups={navLinks} selectedKey={activeScreen} onLinkClick={handleNavClick} styles={navStyles} />
        </div>
        <div className={styles.mainContent}>
          {renderContent()}
        </div>
      </div>
    </div>
  );
}