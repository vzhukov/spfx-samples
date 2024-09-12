import * as React from 'react';
import type { IDataverseComponentProps } from './IDataverseComponentProps';
import { DetailsList, DetailsListLayoutMode, IColumn, Label, MessageBar, MessageBarType, PrimaryButton, SelectionMode, TextField } from '@fluentui/react';
import { IDataverseComponentState } from './IDataverseComponentState';
import styles from './DataverseComponent.module.scss';

export default class DataverseComponent extends React.Component<IDataverseComponentProps, IDataverseComponentState> {

  public constructor(props: IDataverseComponentProps) {
    super(props);
    this.state = {
      entities: [],
      loading: false,
      instanceUrl: this.props.dataverseService.instanceUrl
    };
  }

  public render(): React.ReactElement<IDataverseComponentProps> {
    return (
      <section className={styles.dataverseApiWebPart}>
        <div>
          <TextField label='Dataverse Instance URL'
            value={this.state.instanceUrl}
            onChange={(_, v) => { this.props.dataverseService.instanceUrl = v as string; this.setState({ instanceUrl: v }); }} />
        </div>
        {this.state.error &&
          <div style={{ padding: "20px 0" }}>
            <MessageBar messageBarType={MessageBarType.error} isMultiline={true}>{JSON.stringify(this.state.error)}</MessageBar>
          </div>}
        <div>
          <Label>WhoAmI</Label>
        </div>
        <div>
          <div>
            <PrimaryButton text='Execute WhoAmI' onClick={this.whoAmIHandler.bind(this)} disabled={this.state.loading} />
          </div>
          <div>
            {this.state.whoAmI &&
              <div>
                <TextField label={"Organizationi Id"} value={this.state.whoAmI.OrganizationId} readOnly={true} />
                <TextField label={"User Id"} value={this.state.whoAmI.UserId} readOnly={true} />
              </div>
            }
          </div>
        </div>

        <div style={{ paddingTop: "24px" }}>
          <div>
            <PrimaryButton text='Retrieve Accounts' onClick={this.retrieveAccountsHandler.bind(this)} disabled={this.state.loading} />
          </div>
          {this.state.entities.length > 0 &&
            <div>
              <DetailsList
                items={this.state.entities}
                compact={true}
                columns={this.columns}
                selectionMode={SelectionMode.none}
                getKey={x => x.accountid}
                setKey="none"
                layoutMode={DetailsListLayoutMode.fixedColumns}
                isHeaderVisible={true}
              />
            </div>
          }
        </div>
      </section>
    );
  }

  private columns: IColumn[] = [
    { key: 'name', name: 'Name', fieldName: 'name', minWidth: 150 },
    { key: 'accountnumber', name: 'Number', fieldName: 'accountnumber', minWidth: 150 },
    { key: 'websiteurl', name: 'URL', fieldName: 'websiteurl', minWidth: 200 },
    { key: 'address1_composite', name: 'Address', fieldName: 'address1_composite', minWidth: 250 }
  ];

  private whoAmIHandler(): void {
    this.setState({ loading: true, error: undefined, whoAmI: undefined });
    this.props.dataverseService.whoAmI()
      .then(x => {
        this.setState({ whoAmI: x, loading: false, error: undefined });
      })
      .catch(e => {
        this.setState({ whoAmI: undefined, loading: false, error: e });
      });
  }

  private retrieveAccountsHandler(): void {
    this.setState({ loading: true, error: undefined, entities: [] });

    // Retrieving accounts from Dataverse
    this.props.dataverseService.getEntityList("accounts", ["name", "accountnumber", "websiteurl", "address1_composite"])
      .then(x => {
        this.setState({ entities: x, loading: false, error: undefined });
      })
      .catch(e => {
        this.setState({ entities: [], loading: false, error: e });
      });
  }
}
