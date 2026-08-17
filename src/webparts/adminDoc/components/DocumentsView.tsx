import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './AdminDoc.module.scss';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { DetailsList, SelectionMode, IColumn, DetailsListLayoutMode } from '@fluentui/react/lib/DetailsList';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

interface IDocumentsViewProps {
  context: WebPartContext;
  onCreateNew: () => void;
}

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

export const DocumentsView: React.FC<IDocumentsViewProps> = ({ context, onCreateNew }) => {
  const [documents, setDocuments] = useState<any[]>([]);
  const [filteredDocuments, setFilteredDocuments] = useState<any[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);

  // Filter states
  const [searchText, setSearchText] = useState<string>('');
  const [selectedStatus, setSelectedStatus] = useState<string>('all');
  const [selectedType, setSelectedType] = useState<string>('all');
  const [selectedDept, setSelectedDept] = useState<string>('all');
  const [selectedPivot, setSelectedPivot] = useState<string>('All Documents');

  useEffect(() => {
    fetchSharePointDocuments();
  }, []);

  const fetchSharePointDocuments = async () => {
    setIsLoading(true);
    setErrorMessage(null);

    try {
      const siteUrl = context.pageContext.web.absoluteUrl;
      
      // Use standard GET with proper headers to avoid any OData reader issues
      const response: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('Documents')/items?$orderby=Created desc`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose'
          }
        }
      );

      if (response.ok) {
        const data = await response.json();
        // Accessing the results array safely
        const results = data.d && data.d.results ? data.d.results : (data.value || []);

        const mappedDocs = results.map((item: any) => {
          const fileName = item.FileLeafRef || '';
          const fileExtension = fileName.split('.').pop()?.toLowerCase() || 'word';
          let fileType = 'word';
          if (fileExtension === 'pdf') fileType = 'pdf';
          else if (fileExtension === 'xls' || fileExtension === 'xlsx') fileType = 'excel';

          return {
            key: (item.Id || item.ID || Math.random()).toString(),
            id: item.Id || item.ID || 0,
            docNo: item.Title ? `DOC-${item.Id}` : 'SOP-2026-001',
            title: item.Title || fileName || 'Untitled Document',
            type: item.DocumentType || 'SOP',
            dept: item.Department || 'IT',
            owner: 'Arief Iman Santoso',
            status: item.Status || 'Draft',
            version: item.VersionText || 'v1.0',
            modified: item.Modified ? new Date(item.Modified).toLocaleString() : '08/17/2026 10:30 AM',
            modifiedBy: 'Arief Iman Santoso',
            fileType: fileType
          };
        });

        setDocuments(mappedDocs);
        setFilteredDocuments(mappedDocs);
      } else {
        // Log the actual error to console so we can see what SharePoint is complaining about
        const errorText = await response.text();
        console.error('SharePoint Fetch Error:', errorText);
        setErrorMessage(`Error fetching documents: ${response.statusText}`);
      }
    } catch (error: any) {
      console.error('Error fetching documents:', error);
      setErrorMessage('An unexpected network error occurred.');
    } finally {
      setIsLoading(false);
    }
  };

  // Apply all search and filter conditions
  const applyFilters = (search: string, status: string, type: string, dept: string, pivot: string) => {
    let result = [...documents];

    // Pivot Tab Filter
    if (pivot === 'My Documents') {
      result = result.filter(d => d.owner === 'Arief Iman Santoso');
    } else if (pivot === 'Draft') {
      result = result.filter(d => d.status.toLowerCase() === 'draft');
    } else if (pivot === 'Under Review') {
      result = result.filter(d => d.status.toLowerCase() === 'under review');
    } else if (pivot === 'Pending Approval') {
      result = result.filter(d => d.status.toLowerCase() === 'pending approval');
    } else if (pivot === 'Approved') {
      result = result.filter(d => d.status.toLowerCase() === 'approved');
    } else if (pivot === 'Published') {
      result = result.filter(d => d.status.toLowerCase() === 'published');
    } else if (pivot === 'Archived') {
      result = result.filter(d => d.status.toLowerCase() === 'archived');
    }

    // Search Text
    if (search.trim()) {
      const q = search.toLowerCase();
      result = result.filter(d => d.title.toLowerCase().includes(q) || d.docNo.toLowerCase().includes(q));
    }

    // Status Dropdown
    if (status !== 'all') {
      result = result.filter(d => d.status.toLowerCase().replace(/\s+/g, '') === status.toLowerCase());
    }

    // Type Dropdown
    if (type !== 'all') {
      result = result.filter(d => d.type.toLowerCase() === type.toLowerCase());
    }

    // Dept Dropdown
    if (dept !== 'all') {
      result = result.filter(d => d.dept.toLowerCase() === dept.toLowerCase());
    }

    setFilteredDocuments(result);
  };

  const handleSearchChange = (val: string) => {
    setSearchText(val);
    applyFilters(val, selectedStatus, selectedType, selectedDept, selectedPivot);
  };

  const handleStatusChange = (opt?: IDropdownOption) => {
    const key = opt ? (opt.key as string) : 'all';
    setSelectedStatus(key);
    applyFilters(searchText, key, selectedType, selectedDept, selectedPivot);
  };

  const handleTypeChange = (opt?: IDropdownOption) => {
    const key = opt ? (opt.key as string) : 'all';
    setSelectedType(key);
    applyFilters(searchText, selectedStatus, key, selectedDept, selectedPivot);
  };

  const handleDeptChange = (opt?: IDropdownOption) => {
    const key = opt ? (opt.key as string) : 'all';
    setSelectedDept(key);
    applyFilters(searchText, selectedStatus, selectedType, key, selectedPivot);
  };

  const handlePivotClick = (item?: PivotItem) => {
    if (item && item.props.headerText) {
      const header = item.props.headerText;
      setSelectedPivot(header);
      applyFilters(searchText, selectedStatus, selectedType, selectedDept, header);
    }
  };

  const handleClearFilters = () => {
    setSearchText('');
    setSelectedStatus('all');
    setSelectedType('all');
    setSelectedDept('all');
    setSelectedPivot('All Documents');
    setFilteredDocuments(documents);
  };

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
          <a href="#" className={styles.docLink} onClick={(e) => e.preventDefault()}>
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
        <PrimaryButton text="Create Document" iconProps={{ iconName: 'Add' }} onClick={onCreateNew} />
      </div>

      {/* Filter Bar */}
      <div className={styles.filterBar}>
        <div className={styles.filterItem}>
          <TextField 
            placeholder="Search document..." 
            iconProps={{ iconName: 'Search' }} 
            value={searchText}
            onChange={(e, val) => handleSearchChange(val || '')}
          />
        </div>
        <div className={styles.filterItem}>
          <Dropdown selectedKey={selectedStatus} options={statusOptions} onChange={(e, opt) => handleStatusChange(opt)} />
        </div>
        <div className={styles.filterItem}>
          <Dropdown selectedKey={selectedType} options={typeOptions} onChange={(e, opt) => handleTypeChange(opt)} />
        </div>
        <div className={styles.filterItem}>
          <Dropdown selectedKey={selectedDept} options={deptOptions} onChange={(e, opt) => handleDeptChange(opt)} />
        </div>
        <div className={styles.filterItem}>
          <Dropdown defaultSelectedKey="all" options={ownerOptions} />
        </div>
        <DefaultButton text="Clear" onClick={handleClearFilters} />
      </div>

      {errorMessage && (
        <MessageBar messageBarType={MessageBarType.error} style={{ marginBottom: '16px' }}>
          {errorMessage}
        </MessageBar>
      )}

      {/* Tabbed Container / Table */}
      <div className={styles.tableWrapper}>
        <Pivot 
          aria-label="Document Status Filters" 
          selectedKey={selectedPivot}
          onLinkClick={handlePivotClick}
          style={{ padding: '8px 16px 0 16px', borderBottom: '1px solid #edebe9' }}
        >
          <PivotItem headerText="All Documents" itemIcon="DocumentSet" />
          <PivotItem headerText="My Documents" itemIcon="FollowUser" />
          <PivotItem headerText="Draft" itemIcon="Edit" />
          <PivotItem headerText="Under Review" itemIcon="Chat" />
          <PivotItem headerText="Pending Approval" itemIcon="Clock" />
          <PivotItem headerText="Approved" itemIcon="CheckMark" />
          <PivotItem headerText="Published" itemIcon="Globe" />
          <PivotItem headerText="Archived" itemIcon="Archive" />
        </Pivot>

        {isLoading ? (
          <div style={{ padding: '50px 0', textAlign: 'center' }}>
            <Spinner size={SpinnerSize.large} label="Loading documents from SharePoint..." />
          </div>
        ) : (
          <DetailsList 
            items={filteredDocuments}
            columns={columns}
            selectionMode={SelectionMode.multiple}
            setKey="multiple"
            layoutMode={DetailsListLayoutMode.justified}
            isHeaderVisible={true}
          />
        )}

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
          
          <div>Showing 1 to {filteredDocuments.length} of {documents.length} documents</div>
        </div>
      </div>
    </div>
  );
};