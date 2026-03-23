import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IAnnouncementFormData } from '../models/IAnnouncement';

/** Validation errors keyed by form field name. Empty string means no error. */
export type FormValidationErrors = Partial<Record<keyof IAnnouncementFormData, string>>;

export interface IAnnouncementFormFieldsProps {
  /** Current form field values. */
  formData: IAnnouncementFormData;
  /** Whether a save/create operation is currently in progress. */
  isSubmitting: boolean;
  /** Whether an image upload is currently in progress. */
  isUploadingImage: boolean;
  /** Per-field validation error messages. */
  validationErrors: FormValidationErrors;
  /** Callback when any form field value changes. */
  onFieldChange: (fieldName: keyof IAnnouncementFormData, fieldValue: string) => void;
  /** Callback when the user selects an image file for upload. */
  onImageFileSelected: (selectedFile: File) => void;
  /** Callback when the user submits the form. */
  onSubmit: () => void;
  /** Callback when the user cancels the form. */
  onCancel: () => void;
  /** True when editing an existing announcement, false when adding new. */
  isEditMode: boolean;
  /** SharePoint context needed for image upload service. */
  context: WebPartContext;
}
