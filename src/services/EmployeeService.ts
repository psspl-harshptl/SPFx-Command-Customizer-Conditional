
import { SPFI } from '@pnp/sp';
import { IListItem } from '../models/IListItem';

export type EmployeeStatus = 'Active' | 'Inactive' | 'On Leave';

export class EmployeeService {
  constructor(private readonly _sp: SPFI) {}

  public async getEmployeeById(id: number): Promise<IListItem> {
    return await this._sp.web.lists
      .getByTitle('EmployeeDirectory')
      .items.getById(id)
      .select(
        'Id',
        'Title',
        'EmployeeCode',
        'Department',
        'Designation',
        'EmployeeStatus',
        'JoiningDate',
        'Manager'
      )() as IListItem;
  }

  public async updateEmployeeStatuses(
    employeeIds: number[],
    newStatus: EmployeeStatus
  ): Promise<void> {
    if (employeeIds.length === 0) {
      throw new Error('Select at least one employee.');
    }

    if (employeeIds.some(id => !Number.isInteger(id) || id <= 0)) {
      throw new Error('One or more employee IDs are invalid.');
    }

    if (new Set(employeeIds).size !== employeeIds.length) {
      throw new Error('Duplicate employee IDs were provided.');
    }

    const validStatuses: EmployeeStatus[] = [
      'Active',
      'Inactive',
      'On Leave'
    ];

    if (validStatuses.indexOf(newStatus) === -1) {
      throw new Error('Invalid employee status.');
    }

    const list = this._sp.web.lists.getByTitle('EmployeeDirectory');

    const results = await Promise.all(
      employeeIds.map(id =>
        list.items.getById(id)
          .update({ EmployeeStatus: newStatus })
          .then(
            () => ({ id, success: true }),
            () => ({ id, success: false })
          )
      )
    );

    const failedIds = results
      .filter(result => !result.success)
      .map(result => result.id);

    if (failedIds.length > 0) {
      throw new Error(
        `Failed to update employee item IDs: ${failedIds.join(', ')}. ` +
        'Other employees may already have been updated.'
      );
    }
  }
}
