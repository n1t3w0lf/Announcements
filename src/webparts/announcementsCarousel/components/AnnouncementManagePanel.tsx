import * as React from 'react';
import {
  Panel,
  PanelType,
  DetailsList,
  IColumn,
  SelectionMode,
  ConstrainMode,
  IconButton,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Dialog,
  DialogType,
  DialogFooter,
  DefaultButton
} from '@fluentui/react';
import { IAnnouncementManagePanelProps } from './IAnnouncementManagePanelProps';
import { IAnnouncement, IAnnouncementFormData, createEmptyFormData, CelebrationType, IconPosition } from '../models/IAnnouncement';
import { ImageUploadService } from '../services/ImageUploadService';
import { FormValidationErrors } from './IAnnouncementFormFieldsProps';
import AnnouncementFormFields from './AnnouncementFormFields';
import styles from './AnnouncementManagePanel.module.scss';

/** Internal state for the management panel. */
interface IAnnouncementManagePanelState {
  /** All announcements loaded from the list (not just active ones). */
  announcements: IAnnouncement[];
  /** Whether the list is currently loading. */
  isLoading: boolean;
  /** Error message from the last failed operation. */
  errorMessage: string | null;
  /** Success message from the last completed operation. */
  successMessage: string | null;
  /** Whether the add/edit form is visible. */
  isFormVisible: boolean;
  /** The ID of the announcement being edited, or null if adding new. */
  editingAnnouncementId: number | null;
  /** Current form field values. */
  formData: IAnnouncementFormData;
  /** Whether a create/update operation is in progress. */
  isSubmitting: boolean;
  /** Whether an image upload is in progress. */
  isUploadingImage: boolean;
  /** Per-field validation errors. */
  validationErrors: FormValidationErrors;
  /** The ID of the announcement pending delete confirmation, or null. */
  deleteConfirmAnnouncementId: number | null;
  /** Whether any data mutations have occurred during this panel session. */
  hasDataChanged: boolean;
}

/**
 * Management panel component showing all announcements in a DetailsList
 * with full CRUD support and inline image upload.
 */
export default class AnnouncementManagePanel extends React.Component<IAnnouncementManagePanelProps, IAnnouncementManagePanelState> {

  constructor(props: IAnnouncementManagePanelProps) {
    super(props);

    this.state = {
      announcements: [],
      isLoading: false,
      errorMessage: null,
      successMessage: null,
      isFormVisible: false,
      editingAnnouncementId: null,
      formData: createEmptyFormData(),
      isSubmitting: false,
      isUploadingImage: false,
      validationErrors: {},
      deleteConfirmAnnouncementId: null,
      hasDataChanged: false
    };
  }

  public componentDidUpdate(previousProps: IAnnouncementManagePanelProps): void {
    // Load announcements when the panel opens
    if (this.props.isOpen && !previousProps.isOpen) {
      this.loadAllAnnouncements();
      // Reset state for a fresh panel session
      this.setState({
        isFormVisible: false,
        editingAnnouncementId: null,
        formData: createEmptyFormData(),
        errorMessage: null,
        successMessage: null,
        hasDataChanged: false,
        validationErrors: {},
        deleteConfirmAnnouncementId: null
      });
    }
  }

  /** Load all announcements (not just active) from the data service. */
  private async loadAllAnnouncements(): Promise<void> {
    try {
      this.setState({ isLoading: true, errorMessage: null });
      const allAnnouncements = await this.props.dataService.getAllAnnouncements();
      this.setState({
        announcements: allAnnouncements,
        isLoading: false
      });
    } catch (loadError) {
      console.error('AnnouncementManagePanel: Error loading announcements:', loadError);
      this.setState({
        errorMessage: `Failed to load announcements: ${loadError.message || 'Unknown error'}`,
        isLoading: false
      });
    }
  }

  /** Determine the display status of an announcement based on its date range. */
  private getAnnouncementStatus(announcement: IAnnouncement): { label: string; cssClass: string } {
    const now = new Date();
    const validFromDate = new Date(announcement.ValidFrom);
    const validToDate = new Date(announcement.ValidTo);

    if (now < validFromDate) {
      return { label: 'Scheduled', cssClass: styles.scheduledBadge };
    }
    if (now > validToDate) {
      return { label: 'Expired', cssClass: styles.expiredBadge };
    }
    return { label: 'Active', cssClass: styles.activeBadge };
  }

  /** Format a date for display in the DetailsList. */
  private formatDateForDisplay(dateValue: Date): string {
    const dateObject = new Date(dateValue);
    if (isNaN(dateObject.getTime())) return 'N/A';
    return dateObject.toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'short',
      day: 'numeric'
    });
  }

  /** Build column definitions for the DetailsList. */
  private getListColumns(): IColumn[] {
    return [
      {
        key: 'image',
        name: 'Image',
        minWidth: 70,
        maxWidth: 70,
        onRender: (announcement: IAnnouncement) => {
          if (announcement.AnnouncementImage) {
            return (
              <img
                src={announcement.AnnouncementImage}
                alt={announcement.Title}
                className={styles.listImageThumbnail}
              />
            );
          }
          return <div className={styles.noImagePlaceholder}>No img</div>;
        }
      },
      {
        key: 'title',
        name: 'Title',
        fieldName: 'Title',
        minWidth: 150,
        maxWidth: 250,
        isMultiline: true
      },
      {
        key: 'validFrom',
        name: 'From',
        minWidth: 90,
        maxWidth: 110,
        onRender: (announcement: IAnnouncement) =>
          this.formatDateForDisplay(announcement.ValidFrom)
      },
      {
        key: 'validTo',
        name: 'To',
        minWidth: 90,
        maxWidth: 110,
        onRender: (announcement: IAnnouncement) =>
          this.formatDateForDisplay(announcement.ValidTo)
      },
      {
        key: 'status',
        name: 'Status',
        minWidth: 80,
        maxWidth: 100,
        onRender: (announcement: IAnnouncement) => {
          const announcementStatus = this.getAnnouncementStatus(announcement);
          return (
            <span className={`${styles.statusBadge} ${announcementStatus.cssClass}`}>
              {announcementStatus.label}
            </span>
          );
        }
      },
      {
        key: 'celebrationIcon',
        name: 'Icon',
        minWidth: 50,
        maxWidth: 60,
        onRender: (announcement: IAnnouncement) => {
          if (!announcement.CelebrationIcon || announcement.CelebrationIcon === CelebrationType.None) {
            return <span>-</span>;
          }
          return <span>{announcement.CelebrationIcon}</span>;
        }
      },
      {
        key: 'actions',
        name: 'Actions',
        minWidth: 80,
        maxWidth: 80,
        onRender: (announcement: IAnnouncement) => (
          <div className={styles.actionButtonsContainer}>
            <IconButton
              iconProps={{ iconName: 'Edit' }}
              title="Edit announcement"
              ariaLabel={`Edit ${announcement.Title}`}
              onClick={() => this.handleEditButtonClick(announcement)}
            />
            <IconButton
              iconProps={{ iconName: 'Delete' }}
              title="Delete announcement"
              ariaLabel={`Delete ${announcement.Title}`}
              onClick={() => this.setState({ deleteConfirmAnnouncementId: announcement.Id })}
              styles={{ root: { color: '#a4262c' } }}
            />
          </div>
        )
      }
    ];
  }

  // ---- Add / Edit Handlers ----

  /** Open the form for adding a new announcement. */
  private handleAddNewClick = (): void => {
    this.setState({
      isFormVisible: true,
      editingAnnouncementId: null,
      formData: createEmptyFormData(),
      validationErrors: {},
      errorMessage: null,
      successMessage: null
    });
  };

  /** Open the form pre-populated with an existing announcement for editing. */
  private handleEditButtonClick(announcement: IAnnouncement): void {
    const prepopulatedFormData: IAnnouncementFormData = {
      Title: announcement.Title,
      Description: announcement.Description,
      AnnouncementImage: announcement.AnnouncementImage || '',
      ValidFrom: announcement.ValidFrom ? new Date(announcement.ValidFrom).toISOString() : new Date().toISOString(),
      ValidTo: announcement.ValidTo ? new Date(announcement.ValidTo).toISOString() : new Date().toISOString(),
      CelebrationIcon: announcement.CelebrationIcon || CelebrationType.None,
      CelebrationIconPosition: announcement.CelebrationIconPosition || IconPosition.TopRight,
      CustomIconUrl: announcement.CustomIconUrl || '',
      ImageHeight: String(announcement.ImageHeight || 400),
      ImageWidth: String(announcement.ImageWidth != null ? announcement.ImageWidth : 0),
      ImageFit: announcement.ImageFit || 'cover',
      ImageSizeMode: announcement.ImageSizeMode || 'fit',
      ImageBackgroundType: announcement.ImageBackgroundType || 'gradient',
      ImageBackgroundColor: announcement.ImageBackgroundColor || '#667eea',
      BackgroundColor: announcement.BackgroundColor || '#ffffff',
      TitleColor: announcement.TitleColor || '#333333',
      DescriptionColor: announcement.DescriptionColor || '#333333',
      BorderRadius: String(announcement.BorderRadius != null ? announcement.BorderRadius : 8),
      ShowShadow: announcement.ShowShadow != null ? String(announcement.ShowShadow) : 'true',
      ShowCelebrationIcon: announcement.ShowCelebrationIcon != null ? String(announcement.ShowCelebrationIcon) : 'true',
      CelebrationIconSize: String(announcement.CelebrationIconSize || 60),
      GradientStartColor: announcement.GradientStartColor || '#667eea',
      GradientEndColor: announcement.GradientEndColor || '#764ba2',
      GradientDirection: String(announcement.GradientDirection != null ? announcement.GradientDirection : 135),
      OverlayGradientColor: announcement.OverlayGradientColor || '#000000',
      OverlayOpacity: String(announcement.OverlayOpacity != null ? announcement.OverlayOpacity : 30),
      RedirectUrl: announcement.RedirectUrl || '',
      RedirectTarget: announcement.RedirectTarget || '_self'
    };

    this.setState({
      isFormVisible: true,
      editingAnnouncementId: announcement.Id,
      formData: prepopulatedFormData,
      validationErrors: {},
      errorMessage: null,
      successMessage: null
    });
  }

  /** Update a single form field value. */
  private handleFormFieldChange = (fieldName: keyof IAnnouncementFormData, fieldValue: string): void => {
    this.setState(previousState => ({
      formData: {
        ...previousState.formData,
        [fieldName]: fieldValue
      },
      // Clear validation error for this field when user modifies it
      validationErrors: {
        ...previousState.validationErrors,
        [fieldName]: undefined
      }
    }));
  };

  /** Handle image file selection — upload to Site Assets and set the URL. */
  private handleImageFileSelected = async (selectedFile: File): Promise<void> => {
    try {
      this.setState({ isUploadingImage: true, errorMessage: null });

      const uploadResult = await ImageUploadService.uploadImageToSiteAssets(
        this.props.context,
        selectedFile
      );

      this.setState(previousState => ({
        isUploadingImage: false,
        formData: {
          ...previousState.formData,
          AnnouncementImage: uploadResult.absoluteUrl
        }
      }));
    } catch (uploadError) {
      console.error('AnnouncementManagePanel: Image upload failed:', uploadError);
      this.setState({
        isUploadingImage: false,
        errorMessage: `Image upload failed: ${uploadError.message || 'Unknown error'}`
      });
    }
  };

  /** Validate form data and return errors. */
  private validateFormData(): FormValidationErrors {
    const { formData } = this.state;
    const errors: FormValidationErrors = {};

    if (!formData.Title || formData.Title.trim().length === 0) {
      errors.Title = 'Title is required.';
    } else if (formData.Title.length > 255) {
      errors.Title = 'Title must be 255 characters or less.';
    }

    if (!formData.ValidFrom) {
      errors.ValidFrom = 'Valid From date is required.';
    }

    if (!formData.ValidTo) {
      errors.ValidTo = 'Valid To date is required.';
    }

    if (formData.ValidFrom && formData.ValidTo) {
      const fromDate = new Date(formData.ValidFrom);
      const toDate = new Date(formData.ValidTo);
      if (toDate < fromDate) {
        errors.ValidTo = 'Valid To must be on or after Valid From.';
      }
    }

    if (formData.AnnouncementImage && formData.AnnouncementImage.trim().length > 0) {
      if (!/^https?:\/\//i.test(formData.AnnouncementImage)) {
        errors.AnnouncementImage = 'Image URL must start with http:// or https://';
      }
    }

    if (formData.CelebrationIcon === CelebrationType.Custom) {
      if (!formData.CustomIconUrl || formData.CustomIconUrl.trim().length === 0) {
        errors.CustomIconUrl = 'Custom icon URL is required when celebration icon is set to Custom.';
      }
    }

    return errors;
  }

  /** Submit the form — create or update based on editingAnnouncementId. */
  private handleFormSubmit = async (): Promise<void> => {
    const formErrors = this.validateFormData();
    const hasErrors = Object.values(formErrors).some(errorMsg => errorMsg && errorMsg.length > 0);

    if (hasErrors) {
      this.setState({ validationErrors: formErrors });
      return;
    }

    try {
      this.setState({ isSubmitting: true, errorMessage: null });

      const { formData, editingAnnouncementId } = this.state;

      const parseNum = (val: string): number | undefined => {
        const n = parseInt(val, 10);
        return isNaN(n) ? undefined : n;
      };

      const announcementPayload: Partial<IAnnouncement> = {
        Title: formData.Title.trim(),
        Description: formData.Description,
        AnnouncementImage: formData.AnnouncementImage || undefined,
        ValidFrom: new Date(formData.ValidFrom),
        ValidTo: new Date(formData.ValidTo),
        CelebrationIcon: formData.CelebrationIcon,
        CelebrationIconPosition: formData.CelebrationIconPosition,
        CustomIconUrl: formData.CustomIconUrl || undefined,
        ImageHeight: parseNum(formData.ImageHeight),
        ImageWidth: parseNum(formData.ImageWidth),
        ImageFit: formData.ImageFit ? formData.ImageFit as any : undefined,
        ImageSizeMode: formData.ImageSizeMode as any || 'fit',
        ImageBackgroundType: formData.ImageBackgroundType as any || 'gradient',
        ImageBackgroundColor: formData.ImageBackgroundColor || '#667eea',
        BackgroundColor: formData.BackgroundColor || '#ffffff',
        TitleColor: formData.TitleColor || '#333333',
        DescriptionColor: formData.DescriptionColor || '#333333',
        BorderRadius: parseNum(formData.BorderRadius),
        ShowShadow: formData.ShowShadow === 'true',
        ShowCelebrationIcon: formData.ShowCelebrationIcon === 'true',
        CelebrationIconSize: parseNum(formData.CelebrationIconSize),
        GradientStartColor: formData.GradientStartColor || '#667eea',
        GradientEndColor: formData.GradientEndColor || '#764ba2',
        GradientDirection: parseNum(formData.GradientDirection),
        OverlayGradientColor: formData.OverlayGradientColor || '#000000',
        OverlayOpacity: parseNum(formData.OverlayOpacity),
        RedirectUrl: formData.RedirectUrl || undefined,
        RedirectTarget: (formData.RedirectTarget as '_self' | '_blank') || '_self'
      };

      if (editingAnnouncementId !== null) {
        await this.props.dataService.updateAnnouncement(editingAnnouncementId, announcementPayload);
        this.setState({
          successMessage: `"${formData.Title}" updated successfully.`,
          isFormVisible: false,
          isSubmitting: false,
          hasDataChanged: true
        });
      } else {
        await this.props.dataService.createAnnouncement(announcementPayload);
        this.setState({
          successMessage: `"${formData.Title}" created successfully.`,
          isFormVisible: false,
          isSubmitting: false,
          hasDataChanged: true
        });
      }

      // Reload the list to show updated data
      await this.loadAllAnnouncements();
    } catch (submitError) {
      console.error('AnnouncementManagePanel: Error saving announcement:', submitError);
      this.setState({
        errorMessage: `Failed to save announcement: ${submitError.message || 'Unknown error'}`,
        isSubmitting: false
      });
    }
  };

  /** Cancel the form and hide it. */
  private handleFormCancel = (): void => {
    this.setState({
      isFormVisible: false,
      editingAnnouncementId: null,
      formData: createEmptyFormData(),
      validationErrors: {},
      errorMessage: null
    });
  };

  // ---- Delete Handlers ----

  /** Confirm and execute the deletion of an announcement. */
  private handleDeleteConfirm = async (): Promise<void> => {
    const { deleteConfirmAnnouncementId } = this.state;
    if (deleteConfirmAnnouncementId === null) return;

    try {
      this.setState({ errorMessage: null });
      await this.props.dataService.deleteAnnouncement(deleteConfirmAnnouncementId);

      this.setState({
        deleteConfirmAnnouncementId: null,
        successMessage: 'Announcement deleted successfully.',
        hasDataChanged: true
      });

      await this.loadAllAnnouncements();
    } catch (deleteError) {
      console.error('AnnouncementManagePanel: Error deleting announcement:', deleteError);
      this.setState({
        deleteConfirmAnnouncementId: null,
        errorMessage: `Failed to delete announcement: ${deleteError.message || 'Unknown error'}`
      });
    }
  };

  /** Cancel the delete confirmation dialog. */
  private handleDeleteCancel = (): void => {
    this.setState({ deleteConfirmAnnouncementId: null });
  };

  // ---- Panel Dismiss ----

  /** Handle panel dismissal — notify parent whether data changed. */
  private handlePanelDismiss = (): void => {
    this.props.onDismiss(this.state.hasDataChanged);
  };

  // ---- Render ----

  /** Render the delete confirmation dialog. */
  private renderDeleteConfirmDialog(): JSX.Element | null {
    const { deleteConfirmAnnouncementId, announcements } = this.state;
    if (deleteConfirmAnnouncementId === null) return null;

    const announcementToDelete = announcements.find(
      (announcement) => announcement.Id === deleteConfirmAnnouncementId
    );
    const announcementTitle = announcementToDelete ? announcementToDelete.Title : 'this announcement';

    return (
      <Dialog
        hidden={false}
        onDismiss={this.handleDeleteCancel}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Delete Announcement',
          subText: `Are you sure you want to delete "${announcementTitle}"? This action cannot be undone.`
        }}
        modalProps={{ isBlocking: true }}
      >
        <DialogFooter>
          <PrimaryButton
            text="Delete"
            onClick={this.handleDeleteConfirm}
            styles={{ root: { backgroundColor: '#a4262c', borderColor: '#a4262c' } }}
          />
          <DefaultButton text="Cancel" onClick={this.handleDeleteCancel} />
        </DialogFooter>
      </Dialog>
    );
  }

  /** Render the inline form section (add or edit). */
  private renderFormSection(): JSX.Element | null {
    if (!this.state.isFormVisible) return null;

    const { editingAnnouncementId } = this.state;
    const formHeading = editingAnnouncementId !== null
      ? 'Edit Announcement'
      : 'Add New Announcement';

    return (
      <div className={styles.formSection}>
        <h3 className={styles.formSectionTitle}>{formHeading}</h3>
        <AnnouncementFormFields
          formData={this.state.formData}
          isSubmitting={this.state.isSubmitting}
          isUploadingImage={this.state.isUploadingImage}
          validationErrors={this.state.validationErrors}
          onFieldChange={this.handleFormFieldChange}
          onImageFileSelected={this.handleImageFileSelected}
          onSubmit={this.handleFormSubmit}
          onCancel={this.handleFormCancel}
          isEditMode={editingAnnouncementId !== null}
          context={this.props.context}
        />
      </div>
    );
  }

  /** Render the announcement list or empty state. */
  private renderAnnouncementList(): JSX.Element {
    const { announcements, isLoading } = this.state;

    if (isLoading) {
      return (
        <div className={styles.loadingContainer}>
          <Spinner size={SpinnerSize.large} label="Loading announcements..." />
        </div>
      );
    }

    if (announcements.length === 0) {
      return (
        <div className={styles.emptyStateContainer}>
          <div className={styles.emptyStateIcon}>
            <IconButton
              iconProps={{ iconName: 'Megaphone', styles: { root: { fontSize: 48 } } }}
              disabled
            />
          </div>
          <p className={styles.emptyStateMessage}>
            No announcements yet. Click &quot;Add New Announcement&quot; to create your first one.
          </p>
        </div>
      );
    }

    return (
      <DetailsList
        items={announcements}
        columns={this.getListColumns()}
        selectionMode={SelectionMode.none}
        constrainMode={ConstrainMode.unconstrained}
        isHeaderVisible={true}
        compact={false}
      />
    );
  }

  public render(): React.ReactElement<IAnnouncementManagePanelProps> {
    const { isOpen } = this.props;
    const { errorMessage, successMessage, isFormVisible } = this.state;

    return (
      <>
        <Panel
          isOpen={isOpen}
          type={PanelType.large}
          onDismiss={this.handlePanelDismiss}
          headerText="Manage Announcements"
          isLightDismiss={false}
          closeButtonAriaLabel="Close announcements panel"
        >
          <div className={styles.managePanelContent}>
            {/* Success feedback */}
            {successMessage && (
              <MessageBar
                messageBarType={MessageBarType.success}
                className={styles.feedbackBar}
                onDismiss={() => this.setState({ successMessage: null })}
                dismissButtonAriaLabel="Close success message"
              >
                {successMessage}
              </MessageBar>
            )}

            {/* Error feedback */}
            {errorMessage && (
              <MessageBar
                messageBarType={MessageBarType.error}
                className={styles.feedbackBar}
                onDismiss={() => this.setState({ errorMessage: null })}
                dismissButtonAriaLabel="Close error message"
              >
                {errorMessage}
              </MessageBar>
            )}

            {/* Add New button — hidden when form is already visible */}
            {!isFormVisible && (
              <div className={styles.panelHeader}>
                <PrimaryButton
                  text="Add New Announcement"
                  iconProps={{ iconName: 'Add' }}
                  onClick={this.handleAddNewClick}
                />
              </div>
            )}

            {/* Inline form section */}
            {this.renderFormSection()}

            {/* Announcements list */}
            {this.renderAnnouncementList()}
          </div>
        </Panel>

        {/* Delete confirmation dialog */}
        {this.renderDeleteConfirmDialog()}
      </>
    );
  }
}
