import * as React from 'react';
import {
  TextField,
  DatePicker,
  defaultDatePickerStrings,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  Stack,
  Label,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Toggle,
  Separator
} from '@fluentui/react';
import { IAnnouncementFormFieldsProps } from './IAnnouncementFormFieldsProps';
import { CelebrationType, IconPosition, ImageFitMode, ImageSizeMode, ImageBackgroundType } from '../models/IAnnouncement';
import { CelebrationIconService } from '../services/CelebrationIconService';

/**
 * Reusable form component for creating and editing announcements.
 * Renders all input fields, image upload area, and submit/cancel buttons.
 */
const AnnouncementFormFields: React.FC<IAnnouncementFormFieldsProps> = (props) => {
  const {
    formData,
    isSubmitting,
    isUploadingImage,
    validationErrors,
    onFieldChange,
    onImageFileSelected,
    onSubmit,
    onCancel,
    isEditMode
  } = props;

  /** Hidden file input ref — triggered by the visible upload button. */
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  /** Celebration icon dropdown options built from the CelebrationIconService. */
  const celebrationIconOptions: IDropdownOption[] = CelebrationIconService.getAllIconTypes().map(
    (iconType) => ({
      key: iconType.key,
      text: iconType.text
    })
  );

  /** Image fit mode dropdown options. */
  const imageFitOptions: IDropdownOption[] = [
    { key: ImageFitMode.Cover, text: 'Cover (fill & crop)' },
    { key: ImageFitMode.Contain, text: 'Contain (show full image)' },
    { key: ImageFitMode.Fill, text: 'Fill (stretch to fit)' },
    { key: ImageFitMode.None, text: 'None (original size)' }
  ];

  /** Icon position dropdown options. */
  const iconPositionOptions: IDropdownOption[] = [
    { key: IconPosition.TopLeft, text: 'Top Left' },
    { key: IconPosition.TopRight, text: 'Top Right' },
    { key: IconPosition.BottomLeft, text: 'Bottom Left' },
    { key: IconPosition.BottomRight, text: 'Bottom Right' },
    { key: IconPosition.Center, text: 'Center' }
  ];

  /** Whether the custom icon URL field should be visible. */
  const showCustomIconUrlField = formData.CelebrationIcon === CelebrationType.Custom;

  /** Handle file input change — validates and passes selected file to parent. */
  const handleFileInputChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const selectedFiles = event.target.files;
    if (selectedFiles && selectedFiles.length > 0) {
      onImageFileSelected(selectedFiles[0]);
    }
    // Reset input so the same file can be re-selected if needed
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  /** Convert ISO string to Date object for DatePicker, handling invalid dates. */
  const parseIsoToDate = (isoString: string): Date | undefined => {
    if (!isoString) return undefined;
    const parsedDate = new Date(isoString);
    return isNaN(parsedDate.getTime()) ? undefined : parsedDate;
  };

  /** Handle DatePicker selection — converts Date to ISO string for form state. */
  const handleDateSelected = (
    fieldName: 'ValidFrom' | 'ValidTo',
    selectedDate: Date | null | undefined
  ): void => {
    if (selectedDate) {
      onFieldChange(fieldName, selectedDate.toISOString());
    }
  };

  /** Check if there are any validation errors to display. */
  const hasValidationErrors = Object.values(validationErrors).some(
    (errorMessage) => errorMessage && errorMessage.length > 0
  );

  return (
    <Stack tokens={{ childrenGap: 16 }}>

      {hasValidationErrors && (
        <MessageBar messageBarType={MessageBarType.error}>
          Please fix the errors below before submitting.
        </MessageBar>
      )}

      {/* Title field */}
      <TextField
        label="Title"
        required
        value={formData.Title}
        onChange={(_, newValue) => onFieldChange('Title', newValue || '')}
        errorMessage={validationErrors.Title}
        maxLength={255}
        placeholder="Enter announcement title"
        disabled={isSubmitting}
      />

      {/* Description field */}
      <TextField
        label="Description"
        multiline
        rows={4}
        value={formData.Description}
        onChange={(_, newValue) => onFieldChange('Description', newValue || '')}
        errorMessage={validationErrors.Description}
        placeholder="Enter announcement description (plain text)"
        description="Plain text only. HTML tags will be rendered as-is in the carousel."
        disabled={isSubmitting}
      />

      {/* Redirect URL */}
      <Stack horizontal tokens={{ childrenGap: 16 }}>
        <Stack.Item grow={3}>
          <TextField
            label="Redirect URL (optional)"
            value={formData.RedirectUrl}
            onChange={(_, newValue) => onFieldChange('RedirectUrl', newValue || '')}
            placeholder="https://example.com/page"
            description="If set, clicking the announcement will navigate to this URL."
            disabled={isSubmitting}
          />
        </Stack.Item>
        <Stack.Item grow={1}>
          <Dropdown
            label="Open In"
            options={[
              { key: '_self', text: 'Same Tab' },
              { key: '_blank', text: 'New Tab' }
            ]}
            selectedKey={formData.RedirectTarget || '_self'}
            onChange={(_, option) => {
              if (option) onFieldChange('RedirectTarget', option.key as string);
            }}
            disabled={isSubmitting || !formData.RedirectUrl}
          />
        </Stack.Item>
      </Stack>

      {/* Date fields side by side */}
      <Stack horizontal tokens={{ childrenGap: 16 }}>
        <Stack.Item grow={1}>
          <DatePicker
            label="Valid From"
            isRequired
            value={parseIsoToDate(formData.ValidFrom)}
            onSelectDate={(date) => handleDateSelected('ValidFrom', date)}
            strings={defaultDatePickerStrings}
            placeholder="Select start date"
            disabled={isSubmitting}
          />
          {validationErrors.ValidFrom && (
            <span style={{ color: '#a4262c', fontSize: '12px' }}>{validationErrors.ValidFrom}</span>
          )}
        </Stack.Item>
        <Stack.Item grow={1}>
          <DatePicker
            label="Valid To"
            isRequired
            value={parseIsoToDate(formData.ValidTo)}
            onSelectDate={(date) => handleDateSelected('ValidTo', date)}
            strings={defaultDatePickerStrings}
            placeholder="Select end date"
            disabled={isSubmitting}
          />
          {validationErrors.ValidTo && (
            <span style={{ color: '#a4262c', fontSize: '12px' }}>{validationErrors.ValidTo}</span>
          )}
        </Stack.Item>
      </Stack>

      {/* Image upload section */}
      <Stack tokens={{ childrenGap: 8 }}>
        <Label>Announcement Image</Label>
        <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
          <DefaultButton
            text={isUploadingImage ? 'Uploading...' : 'Upload Image'}
            iconProps={{ iconName: 'Upload' }}
            onClick={() => fileInputRef.current?.click()}
            disabled={isSubmitting || isUploadingImage}
          />
          {isUploadingImage && <Spinner size={SpinnerSize.small} label="Uploading..." />}
          <input
            ref={fileInputRef}
            type="file"
            accept="image/jpeg,image/png,image/gif,image/webp"
            style={{ display: 'none' }}
            onChange={handleFileInputChange}
          />
        </Stack>

        {/* Image preview thumbnail */}
        {formData.AnnouncementImage && (
          <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
            <img
              src={formData.AnnouncementImage}
              alt="Announcement preview"
              style={{
                width: 120,
                height: 80,
                objectFit: 'cover',
                borderRadius: 4,
                border: '1px solid #edebe9'
              }}
            />
            <DefaultButton
              text="Remove"
              iconProps={{ iconName: 'Delete' }}
              onClick={() => onFieldChange('AnnouncementImage', '')}
              disabled={isSubmitting}
            />
          </Stack>
        )}

        {/* Manual URL input as fallback */}
        <TextField
          placeholder="Or paste an image URL directly"
          value={formData.AnnouncementImage}
          onChange={(_, newValue) => onFieldChange('AnnouncementImage', newValue || '')}
          errorMessage={validationErrors.AnnouncementImage}
          disabled={isSubmitting}
        />
      </Stack>

      <Separator>Image Settings</Separator>

      {/* Image size mode selector */}
      <Dropdown
        label="Image Sizing"
        options={[
          { key: ImageSizeMode.Fit, text: 'Use CSS Image Fit' },
          { key: ImageSizeMode.Manual, text: 'Set Manually' }
        ]}
        selectedKey={formData.ImageSizeMode || ImageSizeMode.Fit}
        onChange={(_, option) => {
          if (option) onFieldChange('ImageSizeMode', option.key as string);
        }}
        disabled={isSubmitting}
      />

      {/* CSS fit mode — shown when ImageSizeMode is fit */}
      {formData.ImageSizeMode !== ImageSizeMode.Manual && (
        <Dropdown
          label="Image Fit Mode"
          options={imageFitOptions}
          selectedKey={formData.ImageFit || ImageFitMode.Cover}
          onChange={(_, option) => {
            if (option) onFieldChange('ImageFit', option.key as string);
          }}
          disabled={isSubmitting}
        />
      )}

      {/* Manual dimensions — shown when ImageSizeMode is manual */}
      {formData.ImageSizeMode === ImageSizeMode.Manual && (
        <Stack horizontal tokens={{ childrenGap: 16 }}>
          <Stack.Item grow={1}>
            <TextField
              label="Image Height (px)"
              type="number"
              value={formData.ImageHeight}
              onChange={(_, newValue) => onFieldChange('ImageHeight', newValue || '')}
              min={100}
              max={1200}
              disabled={isSubmitting}
            />
          </Stack.Item>
          <Stack.Item grow={1}>
            <TextField
              label="Image Width (px, 0 = full)"
              type="number"
              value={formData.ImageWidth}
              onChange={(_, newValue) => onFieldChange('ImageWidth', newValue || '')}
              min={0}
              max={2000}
              disabled={isSubmitting}
            />
          </Stack.Item>
        </Stack>
      )}

      {/* Image background type selector */}
      <Dropdown
        label="Image Background"
        options={[
          { key: ImageBackgroundType.Solid, text: 'Solid Color' },
          { key: ImageBackgroundType.Gradient, text: 'Gradient' }
        ]}
        selectedKey={formData.ImageBackgroundType || ImageBackgroundType.Gradient}
        onChange={(_, option) => {
          if (option) onFieldChange('ImageBackgroundType', option.key as string);
        }}
        disabled={isSubmitting}
      />

      {/* Solid color — shown when background type is solid */}
      {formData.ImageBackgroundType === ImageBackgroundType.Solid && (
        <TextField
          label="Background Color"
          value={formData.ImageBackgroundColor}
          onChange={(_, newValue) => onFieldChange('ImageBackgroundColor', newValue || '')}
          placeholder="#667eea"
          disabled={isSubmitting}
        />
      )}

      {/* Gradient settings — shown when background type is gradient */}
      {formData.ImageBackgroundType !== ImageBackgroundType.Solid && (
        <Stack horizontal tokens={{ childrenGap: 16 }}>
          <Stack.Item grow={1}>
            <TextField
              label="Gradient Start Color"
              value={formData.GradientStartColor}
              onChange={(_, newValue) => onFieldChange('GradientStartColor', newValue || '')}
              placeholder="#667eea"
              disabled={isSubmitting}
            />
          </Stack.Item>
          <Stack.Item grow={1}>
            <TextField
              label="Gradient End Color"
              value={formData.GradientEndColor}
              onChange={(_, newValue) => onFieldChange('GradientEndColor', newValue || '')}
              placeholder="#764ba2"
              disabled={isSubmitting}
            />
          </Stack.Item>
          <Stack.Item grow={1}>
            <TextField
              label="Gradient Direction (\u00b0)"
              type="number"
              value={formData.GradientDirection}
              onChange={(_, newValue) => onFieldChange('GradientDirection', newValue || '')}
              min={0}
              max={360}
              disabled={isSubmitting}
            />
          </Stack.Item>
        </Stack>
      )}

      {/* Overlay settings */}
      <Stack horizontal tokens={{ childrenGap: 16 }}>
        <Stack.Item grow={1}>
          <TextField
            label="Overlay Color"
            value={formData.OverlayGradientColor}
            onChange={(_, newValue) => onFieldChange('OverlayGradientColor', newValue || '')}
            placeholder="#000000"
            disabled={isSubmitting}
          />
        </Stack.Item>
        <Stack.Item grow={1}>
          <TextField
            label="Overlay Opacity (%)"
            type="number"
            value={formData.OverlayOpacity}
            onChange={(_, newValue) => onFieldChange('OverlayOpacity', newValue || '')}
            min={0}
            max={100}
            disabled={isSubmitting}
          />
        </Stack.Item>
      </Stack>

      <Separator>Card Appearance</Separator>

      {/* Colors */}
      <Stack horizontal tokens={{ childrenGap: 16 }}>
        <Stack.Item grow={1}>
          <TextField
            label="Background Color"
            value={formData.BackgroundColor}
            onChange={(_, newValue) => onFieldChange('BackgroundColor', newValue || '')}
            placeholder="#ffffff"
            disabled={isSubmitting}
          />
        </Stack.Item>
        <Stack.Item grow={1}>
          <TextField
            label="Title Color"
            value={formData.TitleColor}
            onChange={(_, newValue) => onFieldChange('TitleColor', newValue || '')}
            placeholder="#333333"
            disabled={isSubmitting}
          />
        </Stack.Item>
        <Stack.Item grow={1}>
          <TextField
            label="Description Color"
            value={formData.DescriptionColor}
            onChange={(_, newValue) => onFieldChange('DescriptionColor', newValue || '')}
            placeholder="#333333"
            disabled={isSubmitting}
          />
        </Stack.Item>
      </Stack>

      {/* Layout */}
      <TextField
        label="Border Radius (px)"
        type="number"
        value={formData.BorderRadius}
        onChange={(_, newValue) => onFieldChange('BorderRadius', newValue || '')}
        min={0}
        max={50}
        disabled={isSubmitting}
      />

      <Stack horizontal tokens={{ childrenGap: 16 }}>
        <Stack.Item grow={1}>
          <Toggle
            label="Show Shadow"
            checked={formData.ShowShadow === 'true'}
            onChange={(_, checked) => onFieldChange('ShowShadow', checked ? 'true' : 'false')}
            onText="On"
            offText="Off"
            disabled={isSubmitting}
          />
        </Stack.Item>
      </Stack>

      <Separator>Celebration Icon</Separator>

      {/* Celebration icon settings */}
      <Stack horizontal tokens={{ childrenGap: 16 }}>
        <Stack.Item grow={1}>
          <Toggle
            label="Show Celebration Icon"
            checked={formData.ShowCelebrationIcon === 'true'}
            onChange={(_, checked) => onFieldChange('ShowCelebrationIcon', checked ? 'true' : 'false')}
            onText="On"
            offText="Off"
            disabled={isSubmitting}
          />
        </Stack.Item>
        <Stack.Item grow={1}>
          <TextField
            label="Icon Size (px)"
            type="number"
            value={formData.CelebrationIconSize}
            onChange={(_, newValue) => onFieldChange('CelebrationIconSize', newValue || '')}
            min={20}
            max={200}
            disabled={isSubmitting}
          />
        </Stack.Item>
      </Stack>

      <Stack horizontal tokens={{ childrenGap: 16 }}>
        <Stack.Item grow={1}>
          <Dropdown
            label="Celebration Icon"
            options={celebrationIconOptions}
            selectedKey={formData.CelebrationIcon}
            onChange={(_, option) => {
              if (option) onFieldChange('CelebrationIcon', option.key as string);
            }}
            disabled={isSubmitting}
          />
        </Stack.Item>
        <Stack.Item grow={1}>
          <Dropdown
            label="Icon Position"
            options={iconPositionOptions}
            selectedKey={formData.CelebrationIconPosition}
            onChange={(_, option) => {
              if (option) onFieldChange('CelebrationIconPosition', option.key as string);
            }}
            disabled={isSubmitting}
          />
        </Stack.Item>
      </Stack>

      {/* Custom icon URL — visible only when CelebrationIcon is Custom */}
      {showCustomIconUrlField && (
        <TextField
          label="Custom Icon URL"
          required
          value={formData.CustomIconUrl}
          onChange={(_, newValue) => onFieldChange('CustomIconUrl', newValue || '')}
          errorMessage={validationErrors.CustomIconUrl}
          placeholder="Enter URL for custom celebration icon"
          disabled={isSubmitting}
        />
      )}

      {/* Submit and Cancel buttons */}
      <Stack horizontal tokens={{ childrenGap: 12 }} horizontalAlign="start">
        <PrimaryButton
          text={isEditMode ? 'Save Changes' : 'Add Announcement'}
          onClick={onSubmit}
          disabled={isSubmitting || isUploadingImage}
          iconProps={{ iconName: isEditMode ? 'Save' : 'Add' }}
        />
        <DefaultButton
          text="Cancel"
          onClick={onCancel}
          disabled={isSubmitting}
        />
        {isSubmitting && <Spinner size={SpinnerSize.small} label="Saving..." />}
      </Stack>
    </Stack>
  );
};

export default AnnouncementFormFields;
