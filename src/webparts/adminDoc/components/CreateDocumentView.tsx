import * as React from 'react';
import { useState } from 'react';
import styles from './AdminDoc.module.scss';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Icon } from '@fluentui/react/lib/Icon';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

interface ICreateDocProps {
  context: WebPartContext;
  userDisplayName: string;
  onSuccess: () => void;
}

export const CreateDocumentView: React.FC<ICreateDocProps> = ({ context, userDisplayName, onSuccess }) => {
  const [currentStep, setCurrentStep] = useState<number>(1);
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);

  // Form Data State
  const [docType, setDocType] = useState<string>('SOP');
  const [template, setTemplate] = useState<string>('');
  const [title, setTitle] = useState<string>('');
  const [department, setDepartment] = useState<string>('');

  const [classification, setClassification] = useState<string>('Internal');
  const [confidentiality, setConfidentiality] = useState<string>('Internal Use Only');
  const [effectiveDate, setEffectiveDate] = useState<string>('');
  const [reviewDueDate, setReviewDueDate] = useState<string>('');
  const [tags, setTags] = useState<string>('');
  const [description, setDescription] = useState<string>('');

  const departmentOptions: IDropdownOption[] = [
    { key: 'IT', text: 'IT' },
    { key: 'Legal', text: 'Legal' },
    { key: 'HR', text: 'HR' },
    { key: 'Finance', text: 'Finance' },
    { key: 'Operations', text: 'Operations' },
    { key: 'SCM', text: 'SCM' }
  ];

  const classificationOptions: IDropdownOption[] = [
    { key: 'Internal', text: 'Internal' },
    { key: 'Confidential', text: 'Confidential' },
    { key: 'Public', text: 'Public' }
  ];

  const confidentialityOptions: IDropdownOption[] = [
    { key: 'Internal Use Only', text: 'Internal Use Only' },
    { key: 'Shared with External', text: 'Shared with External' },
    { key: 'Restricted', text: 'Restricted' }
  ];

  const getTemplateOptions = (): IDropdownOption[] => {
    if (docType === 'Policy') {
      return [{ key: 'Policy Template v1.docx', text: 'Policy Template v1.docx' }];
    }
    return [
      { key: 'SOP - Template v2.docx', text: 'SOP - Template v2.docx' },
      { key: 'SOP - Template v3.docx', text: 'SOP - Template v3.docx' }
    ];
  };

  // Handle Real SharePoint List Item Creation with hardcoded ID 6
  // Handle Real SharePoint Document Library Creation & Metadata Update
  const handleCreateDocument = async () => {
    setIsSubmitting(true);
    setSuccessMessage(null);
    setCurrentStep(4);

    try {
      const siteUrl = context.pageContext.web.absoluteUrl;

      // 1. Fetch Form Digest and List Entity Type
      const [digestResponse, typeResponse] = await Promise.all([
        context.spHttpClient.post(`${siteUrl}/_api/contextinfo`, SPHttpClient.configurations.v1, {
          headers: { 'Accept': 'application/json;odata=verbose' }
        }),
        context.spHttpClient.get(`${siteUrl}/_api/web/lists/getbytitle('Documents')?$select=ListItemEntityTypeFullName`, SPHttpClient.configurations.v1)
      ]);

      let digest = '';
      if (digestResponse.ok) {
        const digestData = await digestResponse.json();
        digest = digestData.d ? digestData.d.GetContextWebInformation.FormDigestValue : digestData.FormDigestValue;
      }

      let entityTypeName = 'SP.Data.DocumentItem';
      if (typeResponse.ok) {
        const typeData = await typeResponse.json();
        entityTypeName = typeData.d?.ListItemEntityTypeFullName || typeData.ListItemEntityTypeFullName || entityTypeName;
      }

      // 2. Generate a unique file name
      const sanitizedTitle = (title || 'Untitled_Document').replace(/[^a-zA-Z0-9-_]/g, '_');
      const fileName = `${sanitizedTitle}_${Date.now()}.docx`;

      // 3. Upload file to Document Library RootFolder
      const fileResponse: SPHttpClientResponse = await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('Documents')/RootFolder/Files/add(url='${fileName}',overwrite=true)`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'X-RequestDigest': digest
          },
          body: `[Placeholder content for ${title}]`
        }
      );

      if (!fileResponse.ok) {
        const err = await fileResponse.json();
        throw new Error(err.error?.message?.value || 'Failed to upload file to document library.');
      }

      const fileData = await fileResponse.json();
      const fileItem = fileData.d ? fileData.d : fileData;
      const serverRelativeUrl = fileItem.ServerRelativeUrl;

      if (!serverRelativeUrl) {
        throw new Error('File uploaded, but ServerRelativeUrl was not returned.');
      }

      // 4. Fetch the List Item ID associated with this file
      const itemFieldsResponse: SPHttpClientResponse = await context.spHttpClient.get(
        `${siteUrl}/_api/web/getFileByServerRelativeUrl('${serverRelativeUrl}')/ListItemAllFields?$select=Id`,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=verbose' } }
      );

      if (!itemFieldsResponse.ok) {
        throw new Error('File uploaded, but failed to retrieve associated list item ID.');
      }

      const itemFieldsData = await itemFieldsResponse.json();
      const itemId = itemFieldsData.d ? itemFieldsData.d.Id : itemFieldsData.Id;

      if (!itemId) {
        throw new Error('Could not resolve the List Item ID for the uploaded file.');
      }

      // 5. Prepare metadata payload with matching verbose metadata type
      const metadataPayload = {
        __metadata: { type: entityTypeName },
        Title: title || 'Untitled Document',
        DocDesc: description || '',
        Status: 'Draft',
        DocumentType: docType,
        Department: department,
        VersionText: 'v0.1',
        Classification: classification,
        ConfidentialityLevel: confidentiality,
        EffectiveDate: effectiveDate ? new Date(effectiveDate).toISOString() : null,
        ReviewDueDate: reviewDueDate ? new Date(reviewDueDate).toISOString() : null,
        tags: tags || '',
        OwnerId: 6,
        ReviewersId: { results: [6] },
        ApproversId: { results: [6] }
      };

      // 6. Update metadata using precise matching verbose headers and odata-version 3.0
      const updateResponse = await context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('Documents')/items(${itemId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
            'X-RequestDigest': digest,
            'odata-version': '3.0'
          },
          body: JSON.stringify(metadataPayload)
        }
      );

      if (!updateResponse.ok) {
        const updateErr = await updateResponse.json();
        console.error('Metadata update failed details:', updateErr);
        throw new Error(updateErr.error?.message?.value || 'Failed to update document metadata columns.');
      }

      setSuccessMessage('Document successfully created and metadata saved in the SharePoint Library!');
      setTimeout(() => { onSuccess(); }, 2000);

    } catch (error: any) {
      console.error('Error creating document in library:', error);
      setSuccessMessage(`Error: ${error.message || 'Failed to save document.'}`);
      setIsSubmitting(false);
    }
  };

  const docTypesList = [
    { key: 'SOP', name: 'SOP', icon: 'PageListSolid', desc: 'Standard Operating Procedure' },
    { key: 'Policy', name: 'Policy', icon: 'Certificate', desc: 'Company guidelines & rules' },
    { key: 'Contract', name: 'Contract', icon: 'FileCode', desc: 'Legal agreements & NDAs' },
    { key: 'Form', name: 'Form', icon: 'FormField', desc: 'Request & data collection forms' },
    { key: 'Work Instruction', name: 'Work Instruction', icon: 'CheckList', desc: 'Step-by-step operational tasks' }
  ];

  return (
    <div className={styles.wizardContainer}>
      
      {/* 4-STEP PROGRESS HEADER */}
      <div className={styles.stepProgressBar}>
        <div className={`${styles.stepItem} ${currentStep === 1 ? styles.activeStep : currentStep > 1 ? styles.completedStep : ''}`}>
          <div className={styles.stepCircle}>1</div>
          <div className={styles.stepText}>
            <span className={styles.stepTitle}>Select Type & Template</span>
            <span className={styles.stepSubtitle}>Choose document type and template</span>
          </div>
        </div>

        <div className={`${styles.stepItem} ${currentStep === 2 ? styles.activeStep : currentStep > 2 ? styles.completedStep : ''}`}>
          <div className={styles.stepCircle}>2</div>
          <div className={styles.stepText}>
            <span className={styles.stepTitle}>Document Information</span>
            <span className={styles.stepSubtitle}>Fill mandatory information</span>
          </div>
        </div>

        <div className={`${styles.stepItem} ${currentStep === 3 ? styles.activeStep : currentStep > 3 ? styles.completedStep : ''}`}>
          <div className={styles.stepCircle}>3</div>
          <div className={styles.stepText}>
            <span className={styles.stepTitle}>Review & Confirm</span>
            <span className={styles.stepSubtitle}>Review details and confirm</span>
          </div>
        </div>

        <div className={`${styles.stepItem} ${currentStep === 4 ? styles.activeStep : ''}`}>
          <div className={styles.stepCircle}>4</div>
          <div className={styles.stepText}>
            <span className={styles.stepTitle}>Create Document</span>
            <span className={styles.stepSubtitle}>Document will be generated</span>
          </div>
        </div>
      </div>

      {successMessage && <MessageBar messageBarType={MessageBarType.success}>{successMessage}</MessageBar>}

      {/* BODY SPLIT */}
      <div className={styles.wizardBodyGrid}>
        
        {/* 2/3 FORM PANE */}
        <div className={styles.formPane}>
          
          {currentStep === 1 && (
            <div>
              <h3>Select Document Type</h3>
              <div className={styles.typeCardsGrid}>
                {docTypesList.map((dt) => (
                  <div 
                    key={dt.key}
                    className={`${styles.typeCard} ${docType === dt.key ? styles.selectedTypeCard : ''}`}
                    onClick={() => { setDocType(dt.key); setTemplate(''); }}
                  >
                    <Icon iconName={dt.icon} className={styles.cardIcon} />
                    <div className={styles.cardTitle}>{dt.name}</div>
                    <div className={styles.cardDesc}>{dt.desc}</div>
                  </div>
                ))}
              </div>

              <h3>Select Template</h3>
              <Dropdown 
                placeholder="Select a template..."
                selectedKey={template}
                options={getTemplateOptions()}
                onChange={(e, opt) => opt && setTemplate(opt.key as string)}
              />
            </div>
          )}

          {currentStep === 2 && (
            <div>
              <h3>Fill Mandatory Information</h3>
              <div className={styles.formGrid3Cols}>
                <TextField label="Document Title *" value={title} onChange={(e, val) => setTitle(val || '')} required />
                <Dropdown label="Department *" options={departmentOptions} selectedKey={department} onChange={(e, opt) => opt && setDepartment(opt.key as string)} required />
                
                {/* HARDCODED USER FIELD (DISPLAY ONLY) */}
                <TextField label="Document Owner *" value={userDisplayName} disabled description="Automatically set to your account" />

                <Dropdown label="Classification *" options={classificationOptions} selectedKey={classification} onChange={(e, opt) => opt && setClassification(opt.key as string)} required />
                <Dropdown label="Confidentiality Level" options={confidentialityOptions} selectedKey={confidentiality} onChange={(e, opt) => opt && setConfidentiality(opt.key as string)} />
                <TextField label="Effective Date" type="date" value={effectiveDate} onChange={(e, val) => setEffectiveDate(val || '')} />

                {/* HARDCODED REVIEWERS & APPROVERS */}
                <TextField label="Reviewers *" value={userDisplayName} disabled description="Automatically assigned to you" />
                <TextField label="Approvers *" value={userDisplayName} disabled description="Automatically assigned to you" />

                <TextField label="Review Due Date" type="date" value={reviewDueDate} onChange={(e, val) => setReviewDueDate(val || '')} />

                <div className={styles.formGridFull}>
                  <TextField label="Tags / Keywords" value={tags} onChange={(e, val) => setTags(val || '')} placeholder="Comma-separated tags" />
                </div>
                <div className={styles.formGridFull}>
                  <TextField label="Description" multiline rows={3} value={description} onChange={(e, val) => setDescription(val || '')} />
                </div>
              </div>
            </div>
          )}

          {currentStep === 3 && (
            <div>
              <h3>Review Details & Confirm</h3>
              <div style={{ background: '#faf9f8', padding: '16px', borderRadius: '4px', marginBottom: '20px', lineHeight: '1.6' }}>
                <div><strong>Document Type:</strong> {docType}</div>
                <div><strong>Template:</strong> {template}</div>
                <div><strong>Title:</strong> {title}</div>
                <div><strong>Department:</strong> {department}</div>
                <div><strong>Owner:</strong> {userDisplayName} (ID: 6)</div>
                <div><strong>Reviewers:</strong> {userDisplayName} (ID: 6)</div>
                <div><strong>Approvers:</strong> {userDisplayName} (ID: 6)</div>
                <div><strong>Classification:</strong> {classification} ({confidentiality})</div>
              </div>
            </div>
          )}

          {currentStep === 4 && (
            <div style={{ textAlign: 'center', padding: '40px 0' }}>
              <h3>Generating Document...</h3>
              <p>Saving item and binding user ID 6 to SharePoint...</p>
            </div>
          )}

          {currentStep < 4 && (
            <div className={styles.wizardFooter}>
              {currentStep > 1 ? <DefaultButton text="Back" onClick={() => setCurrentStep(currentStep - 1)} /> : <div />}
              {currentStep < 3 ? (
                <PrimaryButton text="Next" disabled={currentStep === 1 && !template} onClick={() => setCurrentStep(currentStep + 1)} />
              ) : (
                <PrimaryButton text="Confirm & Create" primary disabled={isSubmitting} onClick={handleCreateDocument} />
              )}
            </div>
          )}

        </div>

        {/* 1/3 PREVIEW PANE */}
        <div className={styles.previewPane}>
          <h4>Template Preview</h4>
          <div style={{ fontSize: '12px', color: '#605e5c', marginBottom: '8px' }}>
            <strong>Selected Type:</strong> {docType}<br/>
            <strong>Template:</strong> {template || 'None'}
          </div>
          <div className={styles.previewBox}>[ Live Preview ]</div>
        </div>

      </div>

    </div>
  );
};