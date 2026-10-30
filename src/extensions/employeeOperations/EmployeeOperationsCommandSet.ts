import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  ListViewStateChangedEventArgs,
  RowAccessor
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';

import { getSP } from '../../pnpjsConfig';
import { EmployeeService, EmployeeStatus } from '../../services/EmployeeService';
import { BulkStatusUpdateDialog } from '../../components/BulkStatusUpdateDialog';

export interface IEmployeeOperationsCommandSetProperties {
}

const LOG_SOURCE: string = 'EmployeeOperationsCommandSet';

const CSV_FIELDS: string[] = [
  'ID',
  'Title',
  'EmployeeCode',
  'Department',
  'Designation',
  'EmployeeStatus',
  'JoiningDate',
  'Manager'
];

export default class EmployeeOperationsCommandSet extends BaseListViewCommandSet<IEmployeeOperationsCommandSetProperties> {
  private _employeeService!: EmployeeService;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized EmployeeOperationsCommandSet');

    this._employeeService = new EmployeeService(getSP(this.context));

    this._setCommandVisibility(0);
    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    const selectedRows = event.selectedRows ?? [];

    switch (event.itemId) {
      case 'EMPLOYEE_DETAILS':
        this._showEmployeeDetails(selectedRows).catch(this._handleError);
        break;
      case 'BULK_STATUS_UPDATE':
        this._bulkUpdateStatus(selectedRows).catch(this._handleError);
        break;
      case 'EXPORT_EMPLOYEES_CSV':
        this._exportCsv(selectedRows);
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private _onListViewStateChanged = (_args: ListViewStateChangedEventArgs): void => {
    const selectedCount = this.context.listView.selectedRows?.length ?? 0;
    this._setCommandVisibility(selectedCount);
    this.raiseOnChange();
  }


  private _setCommandVisibility(selectedCount: number): void {
    const detailsCommand = this.tryGetCommand('EMPLOYEE_DETAILS');
    if (detailsCommand) {
      detailsCommand.visible = selectedCount === 1;
    }

    const bulkCommand = this.tryGetCommand('BULK_STATUS_UPDATE');
    if (bulkCommand) {
      bulkCommand.visible = selectedCount >= 2;
    }

    const exportCommand = this.tryGetCommand('EXPORT_EMPLOYEES_CSV');
    if (exportCommand) {
      exportCommand.visible = selectedCount >= 1;
    }
  }


  private async _showEmployeeDetails(selectedRows: readonly RowAccessor[]): Promise<void> {
    if (selectedRows.length !== 1) {
      await Dialog.alert('Select exactly one employee.');
      return;
    }

    const id = Number(selectedRows[0].getValueByName('ID'));
    const employee = await this._employeeService.getEmployeeById(id);

    await Dialog.alert(
      `Name: ${employee.Title}\n` +
      `Employee Code: ${employee.EmployeeCode}\n` +
      `Department: ${employee.Department}\n` +
      `Designation: ${employee.Designation}\n` +
      `Status: ${employee.EmployeeStatus}\n` +
      `Joining Date: ${employee.JoiningDate
        ? new Date(employee.JoiningDate).toLocaleDateString('en-US', {
          month: '2-digit',
          day: '2-digit',
          year: 'numeric'
        })
        : 'Not provided'
      }\n` +
      `Manager: ${employee.Manager}`
    );
  }

  private async _bulkUpdateStatus(selectedRows: readonly RowAccessor[]): Promise<void> {
    const input = await Dialog.prompt('Enter new status (Active, Inactive, On Leave):');
    if (input === undefined) {
      return;
    }

    const ids = selectedRows.map(row => Number(row.getValueByName('ID')));
    await this._employeeService.updateEmployeeStatuses(ids, input.trim() as EmployeeStatus);

    await Dialog.alert(`Updated ${ids.length} employee(s).`);
    window.location.reload();
  }

  private _exportCsv(selectedRows: readonly RowAccessor[]): void {
    const escape = (value: unknown): string => {
      const text = value === undefined || value === null ? '' : String(value);
      return `"${text.replace(/"/g, '""')}"`;
    };

    const lines = [
      CSV_FIELDS.join(','),
      ...selectedRows.map(row =>
        CSV_FIELDS.map(field => escape(row.getValueByName(field))).join(',')
      )
    ];

    const blob = new Blob([lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'Employees.csv';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }

  private _handleError = (error: Error): void => {
    Log.error(LOG_SOURCE, error);
    Dialog.alert(error.message).catch(() => { /* no-op */ });
  }
}
