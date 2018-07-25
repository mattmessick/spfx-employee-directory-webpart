import * as React from 'react';
import { escape, chunk, sortBy } from '@microsoft/sp-lodash-subset';
import {
  IPersonaProps,
  Persona,
  PersonaSize,
  Spinner,
  SpinnerSize,
  SearchBox,
  Overlay
} from 'office-ui-fabric-react';
import { sp } from '@pnp/sp';
import { WebPartTitle, FieldUserRenderer, IPrincipal } from '@pnp/spfx-controls-react';
import styles from './EmployeeDirectory.module.scss';
import { IEmployeeDirectoryProps } from './IEmployeeDirectoryProps';
import { IEmployeeDirectoryState } from './IEmployeeDirectoryState';
import { IUserItem } from './IUserItem';

export default class EmployeeDirectory extends React.Component<IEmployeeDirectoryProps, IEmployeeDirectoryState> {
  
  private _initialState: IEmployeeDirectoryState = {
    users: [],
    search: '',
    loading: true
  };

  constructor(props: IEmployeeDirectoryProps) {
    super(props);

    this.state = this._initialState;

    this._onSearchClear = this._onSearchClear.bind(this);
    this._onSearch = this._onSearch.bind(this);
  }

  public componentWillReceiveProps(props: IEmployeeDirectoryProps) {
    this.resetState();
  }

  public componentDidMount() {
    this.init();
  }

  private init() {
    this.setState({
      loading: true
    }, () => {
      this.fetchUsers()
        .then((users: IUserItem[]): void => {
          this.setState({
            users,
            loading: false
          });
        })
        .catch((error: any) => console.error(error) );
    });
  }

  public render(): React.ReactElement<IEmployeeDirectoryProps> {
    let users = chunk(this.state.users, this.props.columns).map((usersChunk: any[]): JSX.Element => {
      return (
        <div className={ styles.employeeGridRow }>
          {
            usersChunk.map((user: IUserItem): JSX.Element => { 
              return (
                <div className={ styles.employeeGridCol } style={ { width: 100 / this.props.columns + '%' }}>
                  <Persona
                    className={ styles.employee }
                    size={ PersonaSize.size72 }
                    primaryText={ user.Title }
                    secondaryText={ user.JobTitle }
                    tertiaryText={ user.Office }
                    imageUrl={ user.Picture && user.Picture.hasOwnProperty('Url') && user.Picture.Url.match(/\?./) ? user.Picture.Url : '' }
                    onRenderPrimaryText={ this._onRenderPrimaryText.bind(this, user) }
                    onRenderSecondaryText={ this._onRenderSecondaryText }
                    onRenderTertiaryText={ this._onRenderTertiaryText }
                    imageShouldFadeIn={ true }
                  />
                </div>
              );
            })
          }
        </div>
      );
    });
    
    return (
      <div className={ styles.employeeDirectory }>
        <div>
          <WebPartTitle
            displayMode={ this.props.displayMode }
            title={ this.props.title }
            updateProperty={ this.props.updateProperty }
          />
          <SearchBox
            className={ styles.search }
            placeholder="Search Employees"
            onSearch={ this._onSearch }
            onEscape={ this._onSearchClear }
            onClear={ this._onSearchClear }
            value={ this.state.search }
            disabled={ this.state.loading }
          />
          <div className={ styles.employeeGrid }>
            {
              (this.state.users.length > 0) ? 
                users
              :
                (this.state.search && !this.state.loading) && (
                  <div className="ms-textAlignCenter">No employees found.</div>
                )
            }
            {
              (this.state.loading) && (
                <Overlay>
                  <Spinner
                    size={ SpinnerSize.large }
                    ariaLive='assertive'
                  />
                </Overlay>
              )
            }
          </div>
        </div>
      </div>
    );
  }

  private _onRenderPrimaryText (user: IUserItem, props: IPersonaProps): JSX.Element {
    let _user: IPrincipal = {
      id: user.Id.toString(),
      email: user.EMail,
      department: user.Department,
      jobTitle: user.JobTitle,
      sip: user.SipAddress,
      title: user.Title,
      value: '',
      picture: user.Picture && user.Picture.hasOwnProperty('Url') && user.Picture.Url.match(/\?./) ? user.Picture.Url : ''
    };
    
    return (
      <FieldUserRenderer
        users={ [_user] }
        context={ this.props.context }
      />
    );
  }

  private _onRenderSecondaryText(props: IPersonaProps): JSX.Element {
    return ( 
      <div className="ms-fontWeight-bold">
        { props.secondaryText }
      </div>
    );
  }

  private _onRenderTertiaryText = (props: IPersonaProps): JSX.Element => {
    return (
      <div>
        { props.tertiaryText }
      </div>
    );
  }

  private resetState(): void {
    this.setState(this._initialState, () => {
      this.init();
    });
  }

  private _onSearch(search: string = ''): void {
    this.setState({
      search,
      loading: true,
      users: []
    }, () => {
      this.fetchUsers()
        .then((users: IUserItem[]): void => {
          this.setState({
            users,
            loading: false
          });
        })
        .catch((error: any) => console.error(error) );
    });
  }

  private _onSearchClear(): void {
    this._onSearch();
  }

  private fetchUsers(): Promise<IUserItem[]> {
    let filter: string = `ContentTypeId eq '0x010A001C8EF5341A5C27438BA308EA8EBD57AE' and EMail ne null and UserName ne null`;
    let exclude: string[] = this.props.exclude ? this.props.exclude.split('\n') : [];
    let searchProperties: string[] = ['Title', 'Department', 'JobTitle', 'Office', 'EMail'];

    if (this.state.search) {
      filter += ' and (';

      searchProperties.map((property: string, index: number) => {
        filter += `substringof('${this.state.search}', ${property})`;

        if ((index + 1) < searchProperties.length) {
          filter += ' or ';
        }
      });

      filter += ')';
    }

    if (exclude.length > 0) {
      exclude.map((name: string) => {
        filter += ` and EMail ne '${name}'`;
      });
    }

    return sp.web.siteUserInfoList.items
      .select(searchProperties.concat(['Id', 'SipAddress', 'Picture']).join(','))
      .filter(filter)
      .getAll()
      .then((items: any[]) => {
        if (this.state.search) {
          return items;
        }

        return sortBy(items, [this.props.sortBy]);
      });
  }
}
