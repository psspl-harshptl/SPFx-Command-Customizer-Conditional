
import * as React from 'react';
import {
  Dialog,
  DialogType,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  Stack,
  Text
} from '@fluentui/react';

export type EmployeeStatus = 'Active' | 'Inactive' | 'On Leave';

export interface IBulkStatusUpdateDialogProps {
  selectedCount: number;
  onConfirm: (status: EmployeeStatus) => Promise<void>;
  onDismiss: () => void;
}

const statusOptions: IDropdownOption[] = [
  { key: 'Active', text: 'Active' },
  { key: 'Inactive', text: 'Inactive' },
  { key: 'On Leave', text: 'On Leave' }
];

export const BulkStatusUpdateDialog: React.FC<
  IBulkStatusUpdateDialogProps
> = ({ selectedCount, onConfirm, onDismiss }) => {
  const [selectedStatus, setSelectedStatus] =
    React.useState<EmployeeStatus | undefined>();

  const [isSaving, setIsSaving] = React.useState(false);
  const [errorMessage, setErrorMessage] =
    React.useState<string | undefined>();

  const handleContinue = async (): Promise<void> => {
    if (!selectedStatus || isSaving) {
      return;
    }

    setIsSaving(true);
    setErrorMessage(undefined);

    try {
      await onConfirm(selectedStatus);
    } catch (error) {
      setErrorMessage(
        error instanceof Error
          ? error.message
          : 'Unable to update employee statuses.'
      );
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <Dialog
      hidden={false}
      onDismiss={isSaving ? undefined : onDismiss}
      dialogContentProps={{
        type: DialogType.normal,
        title: 'Update Employee Status',
        subText: `You have selected ${selectedCount} employee(s).`
      }}
      modalProps={{
        isBlocking: isSaving
      }}
    >
      <Stack tokens={{ childrenGap: 16 }}>
        <Dropdown
          label="New employee status"
          placeholder="Choose a status"
          options={statusOptions}
          selectedKey={selectedStatus}
          onChange={(_event, option) => {
            if (option) {
              setSelectedStatus(option.key as EmployeeStatus);
              setErrorMessage(undefined);
            }
          }}
          disabled={isSaving}
          required
        />

        {errorMessage && (
          <Text role="alert" style={{ color: '#A4262C' }}>
            {errorMessage}
          </Text>
        )}

        <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }}>
          <DefaultButton
            text="Cancel"
            onClick={onDismiss}
            disabled={isSaving}
          />

          <PrimaryButton
            text={isSaving ? 'Updating...' : 'Continue'}
            onClick={() => {
              void handleContinue();
            }}
            disabled={!selectedStatus || isSaving}
          />
        </Stack>
      </Stack>
    </Dialog>
  );
};
