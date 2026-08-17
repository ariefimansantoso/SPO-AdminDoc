import * as React from 'react';
import styles from './AdminDoc.module.scss';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { DetailsList, SelectionMode, IColumn, DetailsListLayoutMode } from '@fluentui/react/lib/DetailsList';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';

// Dropdown filter options
const statusOptions: IDropdownOption[] = [
  { key: 'all', text: 'All Status' },
  { key: 'draft', text: 'Draft' },
  { key: 'review', text: 'Under Review' },
  { key: 'pending', text: 'Pending Approval' },
  { key: 'approved', text: 'Approved' },
  { key: 'published', text: 'Published' },
  { key: 'archived', text: 'Archived' }
];

const typeOptions: IDropdownOption[] = [
  { key: 'all', text: 'All Document' },
  { key: 'sop', text: 'SOP' },
  { key: 'policy', text: 'Policy' },
  { key: 'contract', text: 'Contract' },
  { key: 'form', text: 'Form' }
];

const deptOptions: IDropdownOption[] = [
  { key: 'all', text: 'All Department' },
  { key: 'it', text: 'IT' },
  { key: 'legal', text: 'Legal' },
  { key: 'hr', text: 'HR' },
  { key: 'finance', text: 'Finance' },
  { key: 'operations', text: 'Operations' },
  { key: 'scm', text: 'SCM' }
];

const ownerOptions: IDropdownOption[] = [
  { key: 'all', text: 'All Owner' },
  { key: 'me', text: 'Me' }
];

// Mock Data
const mockDocuments = [
  { key: '1', docNo: 'SOP-2026-001', title: 'Server Maintenance Procedure', type: 'SOP', dept: 'IT', owner: 'Arief Iman Santoso', status: 'Published', version: 'v1.0', modified: '08/17/2026 10:30 AM', modifiedBy: 'Arief Iman Santoso', fileType: 'word' },
  { key: '2', docNo: 'POL-2026-042', title: 'Employee Code of Conduct', type: 'Policy', dept: 'HR', owner: 'Budi Santoso', status: 'Under Review', version: 'v0.9', modified: '08/16/2026 02:15 PM', modifiedBy: 'System', fileType: 'pdf' },
  { key: '3', docNo: 'FRM-2026-112', title: 'Hardware Request Form', type: 'Form', dept: 'Operations', owner: 'Siti Aminah', status: 'Draft', version: 'v0.1', modified: '08/15/2026 09:00 AM', modifiedBy: 'Siti Aminah', fileType: 'excel' },
];

export const DocumentsView: React.FC = () => {
  const columns: IColumn[] = [
    {
      key: 'docNo', name: 'Document No.', fieldName: 'docNo', minWidth: 130, maxWidth: 160,
      onRender: (item) => {
        let iconName = 'Page'; 
        let iconClass = '';
        if (item.fileType === 'word') { iconName = 'WordDocument'; iconClass = styles.iconWord; }
        if (item.fileType === 'excel') { iconName = 'ExcelDocument'; iconClass = styles.iconExcel; }
        if (item.fileType === 'pdf') { iconName = 'PDFDocument'; iconClass = styles.iconPdf; }
        return (
          <a href="#" className={styles.docLink}>
            <Icon iconName={iconName} className={`${styles.fileIcon} ${iconClass}`} /> {item.docNo}
          </a>
        );
      }
    },
    { key: 'title', name: 'Title', fieldName: 'title', minWidth: 160, maxWidth: 240, isRowHeader: true },
    { key: 'type', name: 'Document Type', fieldName: 'type', minWidth: 100, maxWidth: 120 },
    { key: 'dept', name: 'Department', fieldName: 'dept', minWidth: 90, maxWidth: 110 },
    { 
      key: 'owner', name: 'Owner', fieldName: 'owner', minWidth: 150, maxWidth: 180,
      onRender: (item) => <Persona text={item.owner} size={PersonaSize.size24} />
    },
    { key: 'status', name: 'Status', fieldName: 'status', minWidth: 100, maxWidth: 130 },
    { key: 'version', name: 'Version', fieldName: 'version', minWidth: 50, maxWidth: 70 },
    { 
      key: 'modified', name: 'Modified', fieldName: 'modified', minWidth: 140, maxWidth: 180,
      onRender: (item) => (
        <div className={styles.cellModified}>
          <span>{item.modified}</span>
          <span className={styles.modifiedBy}>by {item.modifiedBy}</span>
        </div>
      )
    },
    { 
      key: 'actions', name: 'Actions', fieldName: 'actions', minWidth: 60, maxWidth: 60,
      onRender: () => (
        <IconButton 
          iconProps={{ iconName: 'MoreVertical' }} 
          menuProps={{ 
            items: [
              { key: 'edit', text: 'Edit', iconProps: { iconName: 'Edit' } },
              { key: 'delete', text: 'Delete', iconProps: { iconName: 'Delete' } },
              { key: 'download', text: 'Download', iconProps: { iconName: 'Download' } }
            ] 
          }} 
        />
      )
    }
  ];

  return (
    <div>
      {/* Header */}
      <div className={styles.pageHeader}>
        <div>
          <h2>Documents</h2>
          <p>Browse, search, and manage all document</p>
        </div>
        <PrimaryButton text="Create Document" iconProps={{ iconName: 'Add' }} />
      </div>

      {/* Filter Bar */}
      <div className={styles.filterBar}>
        <div className={styles.filterItem}>
          <TextField placeholder="Search document..." iconProps={{ iconName: 'Search' }} />
        </div>
        <div className={styles.filterItem}><Dropdown defaultSelectedKey="all" options={statusOptions} /></div>
        <div className={styles.filterItem}><Dropdown defaultSelectedKey="all" options={typeOptions} /></div>
        <div className={styles.filterItem}><Dropdown defaultSelectedKey="all" options={deptOptions} /></div>
        <div className={styles.filterItem}><Dropdown defaultSelectedKey="all" options={ownerOptions} /></div>
        <DefaultButton text="Clear" />
      </div>

      {/* Tabbed Container / Table */}
      <div className={styles.tableWrapper}>
        <Pivot aria-label="Document Status Filters" style={{ padding: '8px 16px 0 16px', borderBottom: '1px solid #edebe9' }}>
          <PivotItem headerText="All Documents" itemIcon="DocumentSet" />
          <PivotItem headerText="My Documents" itemIcon="FollowUser" />
          <PivotItem headerText="Draft" itemIcon="Edit" />
          <PivotItem headerText="Under Review" itemIcon="Chat" />
          <PivotItem headerText="Pending Approval" itemIcon="Clock" />
          <PivotItem headerText="Approved" itemIcon="CheckMark" />
          <PivotItem headerText="Published" itemIcon="Globe" />
          <PivotItem headerText="Archived" itemIcon="Archive" />
        </Pivot>

        <DetailsList
          items={mockDocuments}
          columns={columns}
          selectionMode={SelectionMode.multiple}
          setKey="multiple"
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
        />

        {/* Footer */}
        <div className={styles.tableFooter}>
          <div className={styles.footerLeft}>
            <span>Show</span>
            <Dropdown defaultSelectedKey="10" options={[{key: '10', text: '10'}, {key: '20', text: '20'}, {key: '50', text: '50'}]} styles={{ root: { width: 60 } }} />
            <span>per page</span>
          </div>
          
          <div className={styles.footerMiddle}>
            <IconButton iconProps={{ iconName: 'DoubleChevronLeft' }} title="First Page" />
            <IconButton iconProps={{ iconName: 'ChevronLeft' }} title="Previous Page" />
            <span className={`${styles.pageNumber} ${styles.activePage}`}>1</span>
            <span className={styles.pageNumber}>2</span>
            <span className={styles.pageNumber}>3</span>
            <span className={styles.pageNumber}>4</span>
            <IconButton iconProps={{ iconName: 'ChevronRight' }} title="Next Page" />
            <IconButton iconProps={{ iconName: 'DoubleChevronRight' }} title="Last Page" />
          </div>
          
          <div>Showing 1 to 3 of 1,402 documents</div>
        </div>
      </div>
    </div>
  );
};