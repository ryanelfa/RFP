import { LightningElement, track, wire } from 'lwc';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';
import { createRecord, updateRecord } from 'lightning/uiRecordApi';
import getRFPByName from '@salesforce/apex/RFP2Service.getRFPByName';

import RFP2_OBJECT from '@salesforce/schema/RFP2__c';
import RFP2_NAME_FIELD from '@salesforce/schema/RFP2__c.Name';
import RFP2_CORE_FIELD from '@salesforce/schema/RFP2__c.Contenu__c';

export default class GoogleTestComponent extends LightningElement {
    @track rfpName = '';
    @track rfpCore = '';
    @track existingRFPId = null;
    @track showModal = false;

    handleInputChange(event) {
        const field = event.target.dataset.id;
        if (field === 'rfpName') {
            this.rfpName = event.target.value;
        } else if (field === 'rfpCore') {
            this.rfpCore = event.target.value;
        }
    }

    @wire(getRFPByName, { name: '$rfpName' })
    wiredRFP({ error, data }) {
        if (data && data.length > 0) {
            this.existingRFPId = data[0].Id;
        } else {
            this.existingRFPId = null;
        }
    }

    handleSubmit() {
        if (this.existingRFPId) {
            this.showModal = true;
        } else {
            this.createRFP2Record();
        }
    }

    handleOverwrite() {
        this.showModal = false;
        this.updateRFP2Record(this.existingRFPId);
    }

    handleCreateNew() {
        this.showModal = false;
        this.createRFP2Record();
    }

    handleCancel() {
        this.showModal = false;
        this.showToast('Info', 'Operation cancelled by user.', 'info');
    }

    createRFP2Record() {
        const rfp2Fields = {};
        rfp2Fields[RFP2_NAME_FIELD.fieldApiName] = this.rfpName;
        rfp2Fields[RFP2_CORE_FIELD.fieldApiName] = this.rfpCore;

        const rfp2Record = { apiName: RFP2_OBJECT.objectApiName, fields: rfp2Fields };

        createRecord(rfp2Record)
            .then(() => {
                this.showToast('Success', 'RFP2 details saved successfully', 'success');
            })
            .catch(error => {
                this.handleError('RFP2', error);
            });
    }

    updateRFP2Record(recordId) {
        const rfp2Fields = {};
        rfp2Fields[RFP2_NAME_FIELD.fieldApiName] = this.rfpName;
        rfp2Fields[RFP2_CORE_FIELD.fieldApiName] = this.rfpCore;
        rfp2Fields.Id = recordId;

        const rfp2Record = { fields: rfp2Fields };

        updateRecord(rfp2Record)
            .then(() => {
                this.showToast('Success', 'RFP2 details updated successfully', 'success');
            })
            .catch(error => {
                this.handleError('RFP2', error);
            });
    }

    handleError(recordType, error) {
        let errorMessage = `Error creating ${recordType} record: ${error.body.message}`;

        // Extract detailed error messages if available
        if (error.body && error.body.output && error.body.output.errors && error.body.output.errors.length > 0) {
            const fieldErrors = error.body.output.errors[0].fieldErrors;
            if (fieldErrors) {
                for (const field in fieldErrors) {
                    if (fieldErrors[field].length > 0) {
                        errorMessage = `${fieldErrors[field][0].message}`;
                        break;
                    }
                }
            }

            const pageErrors = error.body.output.errors[0].pageErrors;
            if (pageErrors && pageErrors.length > 0) {
                errorMessage = `${pageErrors[0].message}`;
            }
        }

        console.error(errorMessage);
        this.showToast('Error', errorMessage, 'error');
    }

    showToast(title, message, variant) {
        const evt = new ShowToastEvent({
            title: title,
            message: message,
            variant: variant,
            mode: 'dismissable'
        });
        this.dispatchEvent(evt);
    }
}
