import * as React from 'react';
import styles from './TableView.module.scss';
import { ITableViewProps } from './ITableViewProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as pnp from "sp-pnp-js";
import 'react-bootstrap-table-next/dist/react-bootstrap-table2.min.css';
import BootstrapTable from 'react-bootstrap-table-next';
import 'bootstrap/dist/css/bootstrap.css';
import ToolkitProvider, { Search } from 'react-bootstrap-table2-toolkit';
import paginationFactory from 'react-bootstrap-table2-paginator';
import { Label, ILabelStyles } from 'office-ui-fabric-react/lib/Label';
import { TooltipHost } from 'office-ui-fabric-react/lib/Tooltip';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Breadcrumb, IBreadcrumbItem, IDividerAsProps } from 'office-ui-fabric-react/lib/Breadcrumb';
import { sp } from '@pnp/sp';


function DateFormatter(cell, row, rowIndex) {

  if (cell != null) {
    return (
      <div>{(new Date(cell.toString())).toISOString().slice(0, 10)}</div>
    );
  }
}
function NameFormatter(cell, row, rowIndex) {
  if (row["AssigneeFirstName"] != undefined) {
    if (row != null) {
      return (
        row["AssigneeFirstName"] + " " + row["AssigneeLastName"]
      );
    }
  }
  else {
    if (row != null) {
      return (
        row["Assigneefirstname"] + " " + row["Assigneelastname"]
      );
    }
  }
}

const labelStyles: Partial<ILabelStyles> = {
  root: { margin: '10px 0', selectors: { '&:not(:first-child)': { marginTop: 24 } } },
};

const breadcrumbItem: IBreadcrumbItem[] = [
  { text: 'Home', key: 'Home', href: '#', isCurrentItem: false },
  { text: 'Annual Tax Returns', key: 'ATR', href: '#', isCurrentItem: true },
];

function ActionFormatter(cell, row, rowIndex) {
  if (row != null) {
    return (
      <div><Icon className={styles.eyeIcon} iconName='RedEye' title='Open Case' /></div>
    );
  }
}


export interface ITableViewState {
  getExcelListData: any;
}

export default class TableView extends React.Component<ITableViewProps, ITableViewState> {
  private column;
  private breadcrumbItem: IBreadcrumbItem[];
  constructor(props: ITableViewProps, state: ITableViewState) {
    super(props);
    this.state = {
      getExcelListData: []
    };
  }
  public async componentDidMount() {
    let getcurrentuser;
    await sp.web.siteGroups.getByName('Kelloggadmin').users.get().then(async (resuserdata: any) => {
      resuserdata.map(eachitem => {
        if (eachitem.Email == this.props.context.pageContext.user.email) {
          getcurrentuser = eachitem;
        }
      });
      if (!!getcurrentuser) {
        if (this.props.chooseAtrOrTeq == "ATR") {
          sp.web.lists.getByTitle("ATRMasterDataList").items.orderBy("Modified", false).top(4999).get()
            .then(jsonItems => {
              console.log("get list items", jsonItems);
              this.setState({
                getExcelListData: jsonItems,
              });
            });
        }
        else if (this.props.chooseAtrOrTeq == "TEQ") {
          sp.web.lists.getByTitle("Teqmasterdatalist").items.orderBy("Modified", false).top(4999).get()
            .then(jsonItems => {
              console.log("get list items", jsonItems);
              this.setState({
                getExcelListData: jsonItems,
              });
            });
        }
      }
      else {
        if (this.props.chooseAtrOrTeq == "ATR") {
          sp.web.lists.getByTitle("ATRMasterDataList").items.select('*,AttachmentFiles,Advisoremail/EMail,Advisoremail/Title,Advisoremail/Id').expand("AttachmentFiles,Advisoremail").
            filter(`(Advisoremail/EMail eq '${this.props.context.pageContext.user.email}')`).orderBy("Modified", false).top(4999).get()
            .then(jsonItems => {
              console.log("get list items", jsonItems);
              this.setState({
                getExcelListData: jsonItems,
              });
            });
        }
        else if (this.props.chooseAtrOrTeq == "TEQ") {
          sp.web.lists.getByTitle("Teqmasterdatalist").items.select('*,AttachmentFiles,Advisoremail/EMail,Advisoremail/Title,Advisoremail/Id').expand("AttachmentFiles,Advisoremail").
            filter(`(Advisoremail/EMail eq '${this.props.context.pageContext.user.email}')`).orderBy("Modified", false).top(4999).get()
            .then(jsonItems => {
              console.log("get list items", jsonItems);
              this.setState({
                getExcelListData: jsonItems,
              });
            });
        }
      }
    }).catch(async (err) => {
      if (this.props.chooseAtrOrTeq == "ATR") {
        sp.web.lists.getByTitle("ATRMasterDataList").items.select('*,AttachmentFiles,Advisoremail/EMail,Advisoremail/Title,Advisoremail/Id').top(4999).expand("AttachmentFiles,Advisoremail").
          filter(`(Advisoremail/EMail eq '${this.props.context.pageContext.user.email}')`).orderBy("Modified", false).top(4999).get()
          .then(jsonItems => {
            console.log("get list items", jsonItems);
            this.setState({
              getExcelListData: jsonItems,
            });
          });
      }
      else if (this.props.chooseAtrOrTeq == "TEQ") {
        sp.web.lists.getByTitle("Teqmasterdatalist").items.select('*,AttachmentFiles,Advisoremail/EMail,Advisoremail/Title,Advisoremail/Id').expand("AttachmentFiles,Advisoremail").
          filter(`(Advisoremail/EMail eq '${this.props.context.pageContext.user.email}')`).orderBy("Modified", false).top(4999).get()
          .then(jsonItems => {
            console.log("get list items", jsonItems);
            this.setState({
              getExcelListData: jsonItems,
            });
          });
      }
    });
  }
  public render(): React.ReactElement<ITableViewProps> {
    const { SearchBar, ClearSearchButton } = Search;
    if (this.props.chooseAtrOrTeq === "ATR") {
      this.breadcrumbItem = [
        { text: 'Home', key: 'Home', href: this.props.context.pageContext.web.absoluteUrl, isCurrentItem: false },
        { text: 'Annual Tax Returns', key: 'ATR', href: '#', isCurrentItem: true },
      ];
    }
    else if (this.props.chooseAtrOrTeq === "TEQ") {
      this.breadcrumbItem = [
        { text: 'Home', key: 'Home', href: this.props.context.pageContext.web.absoluteUrl, isCurrentItem: false },
        { text: 'Tax Equalizations', key: 'TEQ', href: '#', isCurrentItem: true },
      ];
    }
    if (this.props.chooseAtrOrTeq === "ATR") {
      this.column =
        [{
          dataField: 'Id',
          text: 'Id',
          sort: true,
          hidden: true
        },
        // {
        //   dataField: 'AssigneeFirstName',
        //   text: 'First Name',
        //   sort: true,
        //   hidden: true
        // },
        // {
        //   dataField: 'AssigneeLastName',
        //   text: 'Full Name',
        //   sort: true,
        //   formatter: NameFormatter,
        // },
        {
          dataField: 'TaxYear',
          text: 'Tax Year',
          sort: true
        },
        {
          dataField: 'TaxPaymentType',
          text: 'Payment Type',
          sort: true
        },
        {
          dataField: 'HostCountry',
          text: 'Country',
          sort: true
        },

        {
          dataField: 'Currency',
          text: 'Currency',
          sort: true
        },
        {
          dataField: 'Amount',
          text: 'Amount',
          sort: true,
          formatter: (cell, row) => {
            if (cell != null) {
              return (
                parseFloat(cell).toFixed(2)
              );
            }
          }
        },

        {
          dataField: 'DueDate',
          text: 'Due On',
          sort: true,
          formatter: DateFormatter,
        },
        {
          dataField: 'casestatus',
          text: 'Status',
          sort: true,
        }, {
          dataField: 'casesubstatus',
          text: 'Sub Status',
          sort: true,
        },
        {
          dataField: 'TaxReturnVersion',
          text: 'Tax Return Version',
          sort: true,
        },
        {
          dataField: 'Id',
          text: 'Action',
          formatter: (cell, row) => {
            return (
              <div><Icon className={styles.eyeIcon} iconName='RedEye' title='Open Case'
                onClick={(e) => {
                  window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/Artcasedetails.aspx?CaseId=` + cell;
                }}

                styles={{
                  root: {
                    marginLeft: "-0.5rem", //Nov-2023: Code Added for Open Case in New Window
                    cursor: "pointer",
                    padding: "1vh",
                    textAlign: "center",
                  },
                }}
              />
                {/* Nov-2023: Code Added for Open Case in New Window - Start */}
                <Icon className={styles.eyeIcon} iconName='OpenInNewTab' title="Open Case in New Tab"
                  onClick={(e) => {
                    window.open(`${this.props.context.pageContext.web.absoluteUrl}/SitePages/Artcasedetails.aspx?CaseId=` + cell, "_blank");
                  }}

                  styles={{
                    root: {
                      //marginLeft: 10,
                      cursor: "pointer",
                      padding: "1vh",
                      textAlign: "center",
                    },
                  }}
                />
                {/* Nov-2023: Code Added for Open Case in New Window - End */}
              </div>
            );
          },
        },
        ];
    }
    else if (this.props.chooseAtrOrTeq === "TEQ") {
      this.column =
        [{
          dataField: 'Id',
          text: 'Id',
          sort: true,
          hidden: true
        },

        {
          dataField: 'Assigneelastname',
          text: 'Full Name',
          sort: true,
          hidden: true
        },
        {
          dataField: 'Assigneefirstname',
          text: 'Name',
          sort: true,
          formatter: (cell, row) => {
            if (row != null) {
              return (
                row["Assigneefirstname"] + " " + row["Assigneelastname"]
              );
            }
          }
        },
        {
          dataField: 'Employeenumber',
          text: 'Employee ID',
          sort: true,

        },
        {
          dataField: 'Taxequalizationcountry',
          text: 'Country',
          sort: true
        },
        {
          dataField: 'Taxyear',
          text: 'Tax Year',
          sort: true
        },
        {
          dataField: 'Currency',
          text: 'Currency',
          sort: true
        },
        {
          dataField: 'TeqAmount',
          text: 'TEQ Amount',
          sort: true,
          formatter: (cell, row) => {
            if (cell != null) {
              return (
                parseFloat(cell).toFixed(2)
              );
            }
          }
        },
        {
          dataField: 'Balancedueorfromcompany',
          text: 'Who Owes',
          sort: true,
          formatter: (cell, row) => {
            if (cell != null) {
              return (

                cell == "Balance Due From Company" ? "Company" : cell == "Balance Due to Company" ? "Employee" : ""
              );
            }
          }
        },
        {
          dataField: 'casestatus',
          text: 'Status',
          sort: true
        },
        {
          dataField: 'casesubstatus',
          text: 'Sub Status',
          sort: true,
        },
        {
          dataField: 'Taxreturnversion',
          text: 'Tax Return Version',
          sort: true,
        },
        {
          dataField: 'Id',
          text: 'Action',
          formatter: (cell, row) => {
            return (
              <div><Icon className={styles.eyeIcon} iconName='RedEye' title='Open Case'
                onClick={(e) => {
                  window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/Teqcasedetails.aspx?CaseId=` + cell;
                }}

                styles={{
                  root: {
                    marginLeft: "-0.5rem", //Nov-2023: Code Added for Open Case in New Window
                    cursor: "pointer",
                    padding: "1vh",
                    textAlign: "center",
                  },
                }}
              />
                {/* Nov-2023: Code Added for Open Case in New Window - Start */}
                <Icon className={styles.eyeIcon} iconName='OpenInNewTab' title="Open Case in New Tab"
                  onClick={(e) => {
                    window.open(`${this.props.context.pageContext.web.absoluteUrl}/SitePages/Teqcasedetails.aspx?CaseId=` + cell, "_blank");
                  }}

                  styles={{
                    root: {
                      //marginLeft: 10,
                      cursor: "pointer",
                      padding: "1vh",
                      textAlign: "center",
                    },
                  }}
                />
                {/* Nov-2023: Code Added for Open Case in New Window - End */}
              </div>
            );
          },
        },
        ];
    }
    else {
      this.column = {
        dataField: 'Id',
        text: 'Id',
        hidden: true
      };
    }

    const pageButtonRenderer = ({
      page,
      active,
      disable,
      title,
      onPageChange
    }) => {
      const handleClick = (e) => {
        e.preventDefault();
        onPageChange(page);
      };
      const activeStyle = {};
      if (active) {
        activeStyle["backgroundColor"] = 'black';
        activeStyle["color"] = 'white';
        activeStyle["borderRadius"] = '0.25rem';
        activeStyle["border"] = "1px solid transparent";
        activeStyle["padding"] = ".375rem .75rem";
        activeStyle["fontSize"] = "1rem";
        activeStyle["lineHeight"] = "1.5";
      } else {
        activeStyle["backgroundColor"] = 'gray';
        activeStyle["color"] = 'black';
        activeStyle["borderRadius"] = '0.25rem';
        activeStyle["border"] = "1px solid transparent";
        activeStyle["padding"] = ".375rem .75rem";
        activeStyle["fontSize"] = "1rem";
        activeStyle["lineHeight"] = "1.5";
      }
      if (typeof page === 'string') {
        activeStyle["backgroundColor"] = 'white';
        activeStyle["color"] = 'black';
        activeStyle["borderRadius"] = '0.25rem';
        activeStyle["border"] = "1px solid transparent";
        activeStyle["padding"] = ".375rem .75rem";
        activeStyle["fontSize"] = "1rem";
        activeStyle["lineHeight"] = "1.5";
        // }
      }
      return (
        <li className="page-item">
          <a href="#" onClick={handleClick} style={activeStyle}>{page}</a>
        </li>
      );
    };

    const options = {
      // hideSizePerPage: true,

      // sizePerPage:6,
      pageButtonRenderer
    };
    return (
      <div className={styles.tableView}>
        <div className={styles.container}>
          {!!this.state.getExcelListData && this.props.chooseAtrOrTeq &&
            <ToolkitProvider
              bootstrap4
              keyField="Id"
              data={this.state.getExcelListData}
              columns={this.column}
              search={{
                searchFormatted: true
              }}
              sort>
              {
                props => (
                  <div>
                    <div className={styles.searchDiv}>
                      <label
                        className={styles.searchLabel}>Filter</label>
                      <SearchBar {...props.searchProps} style={{ width: "40vh" }} placeholder={"Enter Keyword for Filtering"} Label={"Filter"} />
                      {/* <ClearSearchButton {...props.searchProps} /> */}
                    </div>
                    <div style={{ display: "flex" }}>
                      <div className={styles.homeIcon}>
                        <a href="#" ><Icon iconName='Home' style={{ color: "black" }} /></a>
                      </div>
                      <Breadcrumb
                        className={styles.breadcrumb}
                        maxDisplayedItems={4}
                        items={this.breadcrumbItem}
                      />
                    </div>
                    <BootstrapTable
                      responsive
                      bordered={false}
                      hover
                      {...props.baseProps}
                      pagination={paginationFactory(options)}
                    />
                  </div>
                )
              }
            </ToolkitProvider>
          }
        </div>
      </div>
    );

  }
}
