import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import '@fortawesome/fontawesome-free/css/all.min.css';
import './assets/style.css';

import styles from './AciesMetadataSearchWebPart.module.scss';
import * as strings from 'AciesMetadataSearchWebPartStrings';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { has } from '@microsoft/sp-lodash-subset';
import * as JSZip from 'jszip';

export interface IAciesMetadataSearchWebPartProps {
  description: string;
  serchResultsNumber: number;
}

export default class AciesMetadataSearchWebPart extends BaseClientSideWebPart<IAciesMetadataSearchWebPartProps> {
  selectedDownloadItems: { title: string; linkUrl: string; knowledgeCategory: string }[] = [];
  isDownloading: boolean = false; // Add a flag to track the download state
  searchResultsNumber = 0;
  searchesRendered = 0;
  searchTerm = '';
  searchResultsToRender = [];
  baseSite: string = "https://zeomega.sharepoint.com/sites/zehub"
  homeIconHtml = `<img id="home-icon" src="${this.baseSite}/SiteAssets/SPFx/Search/house-32.png" alt="House Icon" style="vertical-align: middle; margin-left: 10px; width: 20px; height: 20px; cursor: pointer;position:relative;bottom:2px;">`;
  groupTermId="9eb1e043-02f0-470e-8131-f17c440392ba";
  TermSets =[
    {
      "name": "knowledgeCategories",
      "id": "ff34dc24-83a1-41c5-bab5-9820c5064768"
    },
    {
      "name": "departments",
      "id": "3d92aa01-f4b6-4cac-944a-4b147e08b134"
    },
    {
      "name": "contentAudiences",
      "id": "1cb53b3d-6bda-477f-8d2e-8213d1db516f"
    },
    {
      "name": "productNames",
      "id": "5b3aa2aa-be35-43a8-ac50-411947efc6b8"
    }
    /* These are the dependant modules that will be filtered based on the selected product name. We call them from a list
    {
      "name": "releaseNumbers",
      "id": "5ac4d136-05d6-4971-ab22-123de0bf62bf"
    },
    {
      "name": "modules",
      "id": "a701789d-87e7-4549-9944-e668a128ed6f"
    }
    */
  ];

  // Define global variables to store the responses
  private productNames: string[] = [];
  private departments: string[] = [];
  private knowledgeCategories: string[] = [];
  private contentAudiences: string[] = [];
  private releaseNumbers: { [key: string]: string[] } = {};
  private modules: { [key: string]: string[] } = {};

  public async render(): Promise<void> {
    this.showLoader();
    await this.renderHtml();
    console.log('Rendered ');
    console.log('LOOK HERE');
    await this.initializeComponents();
    this.hideLoader();
  }

  //--------------------------------------------------- START: Initial Setup ---------------------------------------------------//
  private async renderHtml(): Promise<void> {
    const homePage = "https://zeomega.sharepoint.com/sites/zehub"
    const logoPath = "https://zeomega.sharepoint.com/sites/zehub/SiteAssets/SPFx/Search/logo.png"
    
 
    this.domElement.innerHTML = `
    <style>

    
      #spSiteHeader { display: none !important; }
    </style>
    <section class="${styles.mainBanner} zinnovBanner">

     <!--<div class="dotted-loader ${styles.dotted_loader_container}">
        <div class=" ${styles.dotted_loader}"></div>
      </div>-->
      
      <div class="${styles.Banner}">
        <div class="${styles.ContainerFull}">
          <div class="${styles.Top}">
            <div>
              <!--<a href=${homePage} style="text-decoration: none; cursor: pointer;">
                <img src=${logoPath} alt="Logo">
              </a>-->
            </div>
            <div class="${styles.MenuBtns}">
              <button class="${styles.ContributeBtn}" id="ContributeBtn"><span><i class="fa fa-arrow-right" aria-hidden="true"></i></span> Upload</button> 
              
              <div class="${styles.DropDownContribute}" id="DropDownContribute">
                <div class="${styles.ContentMargin}" id="DropdownContributeContent" style="display: block;">
                  <h3>Upload</h3>
                  
                  <div class="${styles.AddDocument}">
                    <div class="${styles.InputGroup}">
                      <label>File</label>
                      <div class="${styles.d_flex}">
                        <div class="${styles.FileUploadSection}">
                          <button class="${styles.CustomFileUpload}" id="custom-file-upload-main">Upload from System</button>
                          <input type="file" id="file-upload-main" name="file-upload-main" style="display: none;">
                          <div class="${styles.FileName}" id="file-name-main"></div>
                        </div>
                      </div>
                    </div>
                  </div>

                  <!-- <div class="${styles.InputGroup}">
                    <label>File Name*</label>
                    <input type="text" id="fileNameInput">
                  </div> -->

                  <div class="${styles.InputGroup}">
                    <label>Department*</label>
                    <select id="contribute-filter1-input">
                      <option value="" selected></option>
                    </select>  
                    <i class="fa fa-caret-down ${styles.DropIcon}" aria-hidden="true"></i>                
                    <!-- <input type="text" id="practiceInput"> -->
                  </div>

                  <div class="${styles.InputGroup}">
                    <label>Project*</label>
                    <select id="contribute-filter2-input">
                      <option value="" selected></option>
                    </select>
                    <i class="fa fa-caret-down ${styles.DropIcon}" aria-hidden="true"></i>    
                    <!-- <input type="text" id="subPracticeInput"> -->
                  </div>

                  <div class="${styles.InputGroup}">
                    <label>Document Type*</label>
                    <select id="contribute-filter3-input">
                      <option value="" selected></option>
                    </select>
                    <i class="fa fa-caret-down ${styles.DropIcon}" aria-hidden="true"></i>    
                    <i class="fas fa-info-circle ${styles.ClickIcon}" title="Add the relevant product category. Click to see more." id="product-category-info-icon"></i>
                    <div class="dropdown-content ${styles.ClickIcon_Content}" id="product-category-info-dropdown">
                      <div class="${styles.ClickIconInnerContents}">
                        <div class="${styles.ClickIconContents}">
                          <p>We have 6 document types:</p>
                            <ol>
                              <li>Business Development - Documents related to strategies, plans, and activities aimed at growing the business, including pitch decks, proposals, and client engagement letters.</li>
                              <li>Compliance - Documents ensuring adherence to laws, regulations, and company policies, such as NDAs, MSAs, and policy documents.</li>
                              <li>Financial Documents - Documents related to financial transactions and records, including invoices, engagement letters, and financial reports.</li>
                              <li>Internal Team Documents - Documents used within the organization for internal communication and processes, such as meeting minutes (MoM), job descriptions (JD), and internal reports.</li>
                              <li>Media - Documents and assets used for branding and marketing purposes, including brand assets, media kits, and press releases.</li>
                              <li>Process Document - Detailed documents outlining the steps and procedures for executing specific processes, including SOPs, process maps, and workflow diagrams.</li>
                            </ol>
                        </div>
                      </div>
                    </div>
                  </div>

                  <div class="${styles.InputGroup}">
                    <label>Year*</label>
                    <select id="contribute-filter4-input">
                      <option value="" selected style="display:none;">Year</option>
                    </select>
                    <i class="fa fa-caret-down ${styles.DropIcon}" aria-hidden="true"></i>    
                    <i class="fas fa-info-circle ${styles.ClickIcon}" title="Input the publishing year" id="year-info-icon"></i>
                    <div class="dropdown-content ${styles.ClickIcon_Content}" id="year-info-dropdown">
                      <div class="${styles.ClickIconInnerContents}">
                        <div class="${styles.ClickIconContents}">
                            Input the publishing year
                        </div>
                      </div>
                    </div>
                  </div>

                  <div class="${styles.InputGroup}">
                    <label>Client name</label>
                    <select id="contribute-filter5-input">
                      <option value="" selected style="display:none;">Client</option>
                    </select>
                    <i class="fa fa-caret-down ${styles.DropIcon}" aria-hidden="true"></i>    
                    <i class="fas fa-info-circle ${styles.ClickIcon}" title="Add the relevant client name. If not found, go to help and raise a query" id="client-info-icon"></i>
                    <div class="dropdown-content ${styles.ClickIcon_Content}" id="client-info-dropdown">
                      <div class="${styles.ClickIconInnerContents}">
                        <div class="${styles.ClickIconContents}">
                          Add the relevant client name. If not found, go to help and raise a query
                        </div>
                      </div>
                    </div>
                  </div>

                  <div class="${styles.InputGroup}">
                    <label>Client name</label>
                    <select id="contribute-filter6-input">
                      <option value="" selected style="display:none;">Client</option>
                    </select>
                    <i class="fa fa-caret-down ${styles.DropIcon}" aria-hidden="true"></i>    
                    <i class="fas fa-info-circle ${styles.ClickIcon}" title="Add the relevant client name. If not found, go to help and raise a query" id="client-info-icon"></i>
                    <div class="dropdown-content ${styles.ClickIcon_Content}" id="client-info-dropdown">
                      <div class="${styles.ClickIconInnerContents}">
                        <div class="${styles.ClickIconContents}">
                          Add the relevant client name. If not found, go to help and raise a query
                        </div>
                      </div>
                    </div>
                  </div>
                  
                  <div class="${styles.InputGroup}">
                    <label>Keywords</label>
                    <input type="text" id="keywordsInput" placeholder="Photo, Hiring post, LinkedIn, mountains">
                  </div>
                
                  <div class="${styles.confidential}">
                    <div class="${styles.InputGroup} ${styles.InputGroup1}">
                      <label>Confidential</label>
                      <div id="confidentialInput">
                        <label>
                          <input type="radio" name="select-circle" value="yes">
                          Yes
                        </label>
                        <label>
                          <input type="radio" name="select-circle" value="no" checked>
                          No
                        </label>
                      </div>
                    </div>
                    <div class="${styles.InputGroup} ${styles.InputGroup2}">
                      <label>If yes, please specify the email IDs that will have access to this document:</label>
                      <input type="text" id="confidentialDetailsInput" placeholder="IDs separated by semicolon. Eg: user1@zeomega.com; user2@zeomega.com">
                    </div>
                  </div>

                  <div class="${styles.AddDocument}">
                    <div class="${styles.InputGroup}">
                      <label>Add a supporting document, if any*</label>
                      <div class="${styles.d_flex}">
                        <div class="${styles.FileUploadSection}">
                          <button class="${styles.CustomFileUpload}" id="custom-file-upload-1">Upload from System</button>
                          <input type="file" id="file-upload-1" name="file-upload-1" style="display: none;" multiple>
                          <div class="${styles.FileName}" id="file-name-1"></div>
                        </div>
                      </div>
                    </div>
                  </div>

                  <button id="ContributeSubmitBtn" disabled>Submit</button>
                  <a>Your contribution will reflect on the portal once approved.</a>

                </div>

                <div class="${styles.ContributThankYou}" id="ThankYouContentContribute" style="display: none;">
                  <div class="${styles.ThankYouSec}">
                  <svg version="1.1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" x="0px" y="0px"
                      viewBox="0 0 161.2 161.2" enable-background="new 0 0 161.2 161.2" xml:space="preserve">
                      <path class="${styles.path}" fill="none" stroke="#7DB0D5" stroke-miterlimit="10" d="M425.9,52.1L425.9,52.1c-2.2-2.6-6-2.6-8.3-0.1l-42.7,46.2l-14.3-16.4
                      c-2.3-2.7-6.2-2.7-8.6-0.1c-1.9,2.1-2,5.6-0.1,7.7l17.6,20.3c0.2,0.3,0.4,0.6,0.6,0.9c1.8,2,4.4,2.5,6.6,1.4c0.7-0.3,1.4-0.8,2-1.5
                      c0.3-0.3,0.5-0.6,0.7-0.9l46.3-50.1C427.7,57.5,427.7,54.2,425.9,52.1z"/>
                      <circle class="path" fill="none" stroke="#7DB0D5" stroke-width="4" stroke-miterlimit="10" cx="80.6" cy="80.6" r="62.1"/>
                      <polyline class="path" fill="none" stroke="#7DB0D5" stroke-width="6" stroke-linecap="round" stroke-miterlimit="10" points="113,52.8 
                      74.1,108.4 48.2,86.4 "/>

                      <circle class="${styles.spin}" fill="none" stroke="#7DB0D5" stroke-width="4" stroke-miterlimit="10" stroke-dasharray="12.2175,12.2175" cx="80.6" cy="80.6" r="73.9"/>

                  </svg>
                  <h4>Thank you for your contribution</h4>      
                  <p>We are in the process of verifying the document,  and will update it on the portal soon.</p>      
                  </div>

                </div>

              </div>
             
              <button class="${styles.HelpBtn}" id="HelpBtn"><span><i class="fa fa-question" aria-hidden="true"></i></span> Help</button>
                <div class="${styles.DropDownHelp}" id="DropDownHelp">
                  <div class="${styles.ContentMargin}" id="HelpContent">
                    <h3>Help</h3>
                    <div class="${styles.InputGroup}">
                      <label>Please select the category*</label>
                      <div class="${styles.PositionRelative}">
                        <select class="category-select">
                          <option value="" disabled selected>Category</option>
                          <option value="Demo of KM portal">Demo of KM portal</option>
                          <option value="Report an Issue">Report an Issue</option>
                          <option value="Feedback">Feedback</option>
                          <option value="Add/Remove User(s) for access">Add/Remove User(s) for access</option>
                          <option value="Status of Approval">Status of Approval</option>
                          <option value="Other">Other</option>
                        </select>
                        <i class="fa fa-caret-down" aria-hidden="true"></i>
                      </div>
                    </div>

                      <div class="${styles.InputGroup}">
                      <label>Type your query</label>
                      <textarea id="queryTextarea"></textarea>
                    </div>

                    <button id="SubmitBtn">Submit</button>
                    <!-- <a>or email us at <span>knowledge@zeomega.com</span></a> -->
                  </div>

                  <div class="${styles.ContentMargin}" id="ThankYouContent" style="display: none;">
                    <div class="${styles.ThankYouSec}">
                    <svg version="1.1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" x="0px" y="0px"
                        viewBox="0 0 161.2 161.2" enable-background="new 0 0 161.2 161.2" xml:space="preserve">
                        <path class="${styles.path}" fill="none" stroke="#7DB0D5" stroke-miterlimit="10" d="M425.9,52.1L425.9,52.1c-2.2-2.6-6-2.6-8.3-0.1l-42.7,46.2l-14.3-16.4
                        c-2.3-2.7-6.2-2.7-8.6-0.1c-1.9,2.1-2,5.6-0.1,7.7l17.6,20.3c0.2,0.3,0.4,0.6,0.6,0.9c1.8,2,4.4,2.5,6.6,1.4c0.7-0.3,1.4-0.8,2-1.5
                        c0.3-0.3,0.5-0.6,0.7-0.9l46.3-50.1C427.7,57.5,427.7,54.2,425.9,52.1z"/>
                        <circle class="path" fill="none" stroke="#7DB0D5" stroke-width="4" stroke-miterlimit="10" cx="80.6" cy="80.6" r="62.1"/>
                        <polyline class="path" fill="none" stroke="#7DB0D5" stroke-width="6" stroke-linecap="round" stroke-miterlimit="10" points="113,52.8 
                        74.1,108.4 48.2,86.4 "/>

                        <circle class="${styles.spin}" fill="none" stroke="#7DB0D5" stroke-width="4" stroke-miterlimit="10" stroke-dasharray="12.2175,12.2175" cx="80.6" cy="80.6" r="73.9"/>

                    </svg>
                    <h4>Thank you for your contribution</h4>      
                    <p>We are in the process of verifying the document,  and will update it on the portal soon.</p>      
                    </div>

                  </div>
              </div>
            </div>
          </div>

            <div class="${styles.Bottom}">
              <h2>Welcome to ZeHub</h2>
              <div class="${styles.SearchGroup}">
                <div class="${styles.SearchInput}">
                  <input type="search" placeholder="Search..." id="search-input">
                  <div class="${styles.SearchBtn}" id="search-btn"><i class="fa fa-search" aria-hidden="true"></i></div>
                </div>
                
                <div class="${styles.filterBox}" id="filter-box">
                  <div class="${styles.Row}">
                   
                    <div class="${styles.Col7}">
                      <div class="${styles.RightBox}">
                        <h5>
                          <input type="checkbox" class="${styles.Circle}" name="checkboxGroup" id="filenamesonlycheckbox">
                          <label for="filenamesonlycheckbox" class="${styles.CircleLabel}"></label>
                          Search in File Names Only
                        </h5>
                        
                        <div class="${styles.Row}">

                          <div class="${styles.Col6}">
                            <div class="dropdown ${styles.InputGroup} ${styles.InputMultiSelect}">
                              <select class="category-select" id="filter1-select" multiple>
                              </select>
                              <div class="dropdown-btn  ${styles.dropdown_btn}">
                                <span>Department</span>
                              </div>
                              <div id = "filter1-dropdown-select" class="dropdown-content ${styles.dropdown_content}"></div>
                              <div class="selected-items ${styles.selected_items}"></div>
                            </div>
                          </div>

                          <div class="${styles.Col6}">
                            <div class="${styles.InputGroup}">
                              <select class="category-select" id="filter2-select">
                                <option value="" selected hidden>Product Name</option>
                              </select>
                              <i class="fa fa-caret-down" aria-hidden="true"></i>
                            </div>
                          </div>

                          <!--
                          <div class="${styles.Col6}">
                            <div class="dropdown ${styles.InputGroup} ${styles.InputMultiSelect}">
                              <select class="category-select" id="filter2-select">
                              </select>
                              <div class="dropdown-btn  ${styles.dropdown_btn}">
                                <span>Product Name</span>
                              </div>
                              <div id = "filter2-dropdown-select" class="dropdown-content ${styles.dropdown_content}"></div>
                              <div class="selected-items ${styles.selected_items}"></div>
                            </div>
                          </div>
                          -->

                          <div class="${styles.Col6}">
                            <div class="dropdown ${styles.InputGroup} ${styles.InputMultiSelect}">
                              <select class="category-select" id="filter3-select" multiple>
                              </select>
                              <div class="dropdown-btn  ${styles.dropdown_btn}">
                                <span>Knowledge Category</span>
                              </div>
                              <div id = "filter3-dropdown-select" class="dropdown-content ${styles.dropdown_content}"></div>
                              <div class="selected-items ${styles.selected_items}"></div>
                            </div>
                          </div>

                          <div class="${styles.Col6}">
                            <div class="dropdown ${styles.InputGroup} ${styles.InputMultiSelect}">
                              <select class="category-select" id="filter4-select" multiple>
                              </select>
                              <div class="dropdown-btn  ${styles.dropdown_btn}">
                                <span>Content Audience</span>
                              </div>
                              <div id = "filter4-dropdown-select" class="dropdown-content ${styles.dropdown_content}"></div>
                              <div class="selected-items ${styles.selected_items}"></div>
                            </div>
                          </div>
                                
                          <div class="${styles.Col6}">
                            <div class="dropdown ${styles.InputGroup} ${styles.InputMultiSelect}">
                              <select class="category-select" id="filter5-select" multiple>
                              </select>
                              <div class="dropdown-btn  ${styles.dropdown_btn}">
                                <span>Release Number</span>
                              </div>
                              <div id = "filter5-dropdown-select" class="dropdown-content ${styles.dropdown_content}"></div>
                              <div class="selected-items ${styles.selected_items}"></div>
                            </div>
                          </div>

                          <div class="${styles.Col6}">
                            <div class="dropdown ${styles.InputGroup} ${styles.InputMultiSelect}">
                              <select class="category-select" id="filter6-select" multiple>
                              </select>
                              <div class="dropdown-btn  ${styles.dropdown_btn}">
                                <span>Module</span>
                              </div>
                              <div id = "filter6-dropdown-select" class="dropdown-content ${styles.dropdown_content}"></div>
                              <div class="selected-items ${styles.selected_items}"></div>
                            </div>
                          </div>
                                
                          <div class="${styles.Col12}">
                              <div class="${styles.InputGroup}">
                                <input type="text" id="keyword-input" placeholder="Additional Keywords" style=" border: none;border-bottom: 1px solid #ccc; color: #222; outline: 0; padding: 8px 10px; width: 100%;">
                              </div>
                          </div>

                        </div>  
                      </div>
                    </div>      
                  </div>
                </div>
              </div>         
              
              <p>Discover insights, answers, and resources swiftly and intelligently.</p>
           
              <!-- This is where "search-result-banner-filter" once was-->

              </div> 
            </div>
          </div>
        </div>

      <div class="${styles.SearchResults}" id="search-results">
        <div class="${styles.ContainerFull}">
            <h2 id="search-results-heading">
              ${this.homeIconHtml} 
              Search Results
            </h2>
            <div class="${styles.SearchTag}" id="search-tags">
            </div>
            
            <!--      
            <div class="${styles.DownloadBtn}">
              <button id="downloadAsCsvBtn">Download as CSV</button>
              <button id = "downloadAsZipBtn">Download as ZIP</button>
            </div>
            -->
                    
            <div id="search-result"></div>
            <button id="load-more-btn" class="${styles.LoadMoreBtn}">Load More</button>
        </div>
      </div>
    </section>
    
    
    `;
  }

  private showLoader(): void {
    document.body.classList.add(styles.dotted_loader_no_scroll);
    const loader = document.createElement('div');
    loader.className = `dotted-loader ${styles.dotted_loader_container}`;
    loader.innerHTML = `<div class="${styles.dotted_loader}"></div>`;
    document.body.appendChild(loader);
    console.log('Loader shown');
  }

  private hideLoader(): void {
    document.body.classList.remove(styles.dotted_loader_no_scroll);
    const loader = document.querySelector(`.${styles.dotted_loader_container}`);
    if (loader) {
      document.body.removeChild(loader);
    }
  }

  // Initialize all the necessary components such as data, event listeners, etc. after the render
  private async initializeComponents(): Promise<void> {
    //Filter 1 = Departments
    //Filter 2 = Product Name
    //Filter 3 = Knowledge Category
    //Filter 4 = Content Audience
    //Filter 5 = Release Number
    //Filter 6 = Module

    //--------------------------------------------------------------------------------------------------------//

    //Part 1 - Fetch all the filter data from the lists and populate the dropdowns with them
    await this.fetchFromTermstore();

    /*These methods are replaced by the fetchFromTermstore method
    await this.fetchDepartments();
    await this.fetchProductNames();
    await this.fetchKnowledgeCategory();
    await this.fetchContentAudience();
    */
    await this.fetchReleaseNumbers();
    await this.fetchModules();
    

    this.populateDepartmentOptions();
    this.populateProductNames();
    this.populateKnowledgeCategory();
    this.populateContentAudience();
    //These are the dependent dropdowns that will be filtered based on the selected product name
    this.populateModules(''); //Initially, no product name is selected, so show all modules
    this.populateReleaseNumbers(''); //Initially, no product name is selected, so show all release numbers

    //--------------------------------------------------------------------------------------------------------//

    //Part 2 - Add event listeners to the filter box and filter dropdowns
    this.addFilterBoxEventListeners(); //Event handler to show and hide the filterbox based on clicks
    this.addProductChangeListener(); //Event handler to update the dependent dropdowns of modules and release numbers based on the selected product name
    this.addSelectChangeListener(); //This sets the selected options in the filter dropdowns and displays them in a dark color
    this.addMultiSelectFeaturesToDropdowns(); //ensures that the dropdown content is populated with options, handles the toggling of dropdown visibility, manages the selection of options, and closes dropdowns when clicking outside of them
    
    //--------------------------------------------------------------------------------------------------------//

    //Part 3 - Search-related event listeners
    this.addPerformSearchListeners(); //adds event listeners to the search button and the enter key to perform the search
    //Search tags are created when the search results are rendered
    this.addTagsCloseIconListeners(); // when search is performed, the selected filters are displayed as tags. This event listener will remove the tag when the close icon is clicked

    //--------------------------------------------------------------------------------------------------------//

    //Part 4 - Add event listeners for the contribute feature
    this.addSupportingFileUploadListeners();
    this.addMainFileUploadListeners();
    this.addDropDownContributeListeners();
    this.addContributeSubmitButtonListener();
    this.addEventHandlersForValidation();
    this.addInfoIconEventListeners();
    
    //--------------------------------------------------------------------------------------------------------//

    //Part 5 - Add event listeners for the help feature
    this.DropHelpButtonListener();
    
    //--------------------------------------------------------------------------------------------------------//

  }

  private addPerformSearchListeners(): void {
    const searchInput = this.domElement.querySelector('#search-input') as HTMLInputElement;
    const filterBox = this.domElement.querySelector('#filter-box') as HTMLDivElement;
    const searchBtn = this.domElement.querySelector('#search-btn') as HTMLDivElement;
  
    searchInput.addEventListener('click', (event) => {
      event.stopPropagation(); // Prevent the click event from bubbling up to the document
      filterBox.style.display = 'block';
    });

    // Show search results when Enter key is pressed inside the search input
    searchInput.addEventListener('keydown', async (event) => {
      if (event.key === 'Enter') {
        await this.handleSearchResults();
      }
    });
  
    document.addEventListener('click', (event) => {
      const target = event.target as Node;
      if (!searchInput.contains(target) && !filterBox.contains(target)) {
        filterBox.style.display = 'none';
      }
    });
  
    // Ensure clicks inside the filter box do not close it
    filterBox.addEventListener('click', (event) => {
      event.stopPropagation();
    });
  
    // Show search results and search result banner filter when the search button is clicked
    searchBtn.addEventListener('click', async () => {
      await this.handleSearchResults();
    });
  }
  //---------------------------------------------------- END: Initial Setup ----------------------------------------------------//

  

  //--------------------------------------- START: Fetch Data and Populate the Dropdowns ---------------------------------------//
  private async fetchFromTermstore() {
    // Fetch the request digest
    const digestResponse = await fetch(`${this.context.pageContext.web.absoluteUrl}/_api/contextinfo`, {
      method: 'POST',
      headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose'
      }
    });

    if (!digestResponse.ok) {
        throw new Error(`HTTP error! status: ${digestResponse.status}`);
    }

    const digestData = await digestResponse.json();
    const requestDigest = digestData.d.GetContextWebInformation.FormDigestValue;

    for (const termset of this.TermSets) {
      const termsetId = termset.id;
      const termsetName = termset.name;

      //These are the dependant dropdowns we should not fetch from the termstore
      if (termsetName === 'releaseNumbers' || termsetName === 'modules') {
        break;
      }

      // Send a post request to the list
      const listApiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/v2.1/termStore/groups/${this.groupTermId}/sets/${termsetId}/terms`;

      const response = await fetch(listApiUrl, {
          method: 'GET',
          headers: {
              'Accept': 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose',
              'X-RequestDigest': requestDigest
          }
      });

      const allTermsUnderTermset = await response.json();

      allTermsUnderTermset.value.forEach(termObject => {
        termObject.labels.forEach(label => {
          if (!Array.isArray(this[termsetName])) {
            this[termsetName] = [];
          }
          this[termsetName].push(label.name);
        });
      });
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
    };

    console.log('Departments:', this.departments);
    console.log('Product Names:', this.productNames);
    console.log('Knowledge Categories:', this.knowledgeCategories);
    console.log('Content Audiences:', this.contentAudiences);
    console.log('Release Numbers:', this.releaseNumbers);
    console.log('Modules:', this.modules);
  }
  
  /* This entire section is made was used to fetch data from the lists. This is now replaced by the termstore fetch. Only the dependant dropdowns are fetched from the lists
  private async fetchDepartments(): Promise<void> {
    const response = await this.context.spHttpClient.get(
      `${this.baseSite}/_api/web/lists/getbytitle('Search_Departments')/items`,
      SPHttpClient.configurations.v1
    );
    const data = await response.json();
    this.departments = data.value.map(item => item.Title);
    console.log('Departments:', this.departments);
  }

  private async fetchProductNames(): Promise<void> {
    const response = await this.context.spHttpClient.get(
      `${this.baseSite}/_api/web/lists/getbytitle('Search_ProductNames')/items`,
      SPHttpClient.configurations.v1
    );
    const data = await response.json();
    this.productNames = data.value.map(item => item.Title);
    console.log('Product Names:', this.productNames);
  }

  private async fetchKnowledgeCategory(): Promise<void> {
    const response = await this.context.spHttpClient.get(
      `${this.baseSite}/_api/web/lists/getbytitle('Search_KnowledgeCategories')/items`,
      SPHttpClient.configurations.v1
    );
    const data = await response.json();
    this.knowledgeCategories = data.value.map(item => item.Title);
    console.log('Knowledge Categories:', this.knowledgeCategories);
  }

  private async fetchContentAudience(): Promise<void> {
    const response = await this.context.spHttpClient.get(
      `${this.baseSite}/_api/web/lists/getbytitle('Search_ContentAudience')/items`,
      SPHttpClient.configurations.v1
    );
    const data = await response.json();
    this.contentAudiences = data.value.map(item => item.Title);
    console.log('Content Audiences:', this.contentAudiences);
  }
  */

  private async fetchReleaseNumbers(): Promise<void> {
    const response = await this.context.spHttpClient.get(
      `${this.baseSite}/_api/web/lists/getbytitle('Search_ReleaseNumbers')/items?$select=Title,ReleaseNumbers`,
      SPHttpClient.configurations.v1
    );
    const data = await response.json();
    data.value.forEach(item => {
      if (!this.releaseNumbers[item.Title]) {
        this.releaseNumbers[item.Title] = [];
      }
      this.releaseNumbers[item.Title].push(item.ReleaseNumbers);
    });
    console.log('Release Numbers:', this.releaseNumbers);
  }

  private async fetchModules(): Promise<void> {
    const response = await this.context.spHttpClient.get(
      `${this.baseSite}/_api/web/lists/getbytitle('Search_Modules')/items?$select=Title,Modules`,
      SPHttpClient.configurations.v1
    );
    const data = await response.json();
    data.value.forEach(item => {
      if (!this.modules[item.Title]) {
        this.modules[item.Title] = [];
      }
      this.modules[item.Title].push(item.Modules);
    });
    console.log('Modules:', this.modules);
  }
  

  private populateDepartmentOptions(): void {
    const filter1Select = this.domElement.querySelector('#filter1-select') as HTMLSelectElement;
    //console.log("filter1Select", filter1Select);
    this.departments.forEach(item => {
      const option = document.createElement('option');
      option.value = item;
      option.text = item;
      filter1Select.appendChild(option);
    });

    const contributeFilter1Input = this.domElement.querySelector('#contribute-filter1-input') as HTMLSelectElement;
    contributeFilter1Input.innerHTML = '<option value="" selected style="display:none;">Department</option>';
    //console.log("filter1Select", practiceInputSelect);
      this.departments.forEach(item => {
        const option = document.createElement('option');
        option.value = item;
        option.text = item;
        contributeFilter1Input.appendChild(option);
    });
  }

  private populateProductNames(): void {
    const filter2Select = this.domElement.querySelector('#filter2-select') as HTMLSelectElement;
    //console.log("filter1Select", filter1Select);
    this.productNames.forEach(item => {
      const option = document.createElement('option');
      option.value = item;
      option.text = item;
      filter2Select.appendChild(option);
    });

    const contributeFilter2Input = this.domElement.querySelector('#contribute-filter2-input') as HTMLSelectElement;
    contributeFilter2Input.innerHTML = '<option value="" selected style="display:none;">Product Name</option>';
    //console.log("filter1Select", practiceInputSelect);
      this.productNames.forEach(item => {
        const option = document.createElement('option');
        option.value = item;
        option.text = item;
        contributeFilter2Input.appendChild(option);
    });
  }

  private populateKnowledgeCategory(): void {
    const filter3Select = this.domElement.querySelector('#filter3-select') as HTMLSelectElement;
    //console.log("filter1Select", filter1Select);
    this.knowledgeCategories.forEach(item => {
      const option = document.createElement('option');
      option.value = item;
      option.text = item;
      filter3Select.appendChild(option);
    });

    const contributeFilter3Input = this.domElement.querySelector('#contribute-filter3-input') as HTMLSelectElement;
    contributeFilter3Input.innerHTML = '<option value="" selected style="display:none;">Knowledge Category</option>';
    //console.log("filter1Select", practiceInputSelect);
      this.knowledgeCategories.forEach(item => {
        const option = document.createElement('option');
        option.value = item;
        option.text = item;
        contributeFilter3Input.appendChild(option);
    });
  }

  private populateContentAudience(): void {
    const filter4Select = this.domElement.querySelector('#filter4-select') as HTMLSelectElement;
    //console.log("filter1Select", filter1Select);
    this.contentAudiences.forEach(item => {
      const option = document.createElement('option');
      option.value = item;
      option.text = item;
      filter4Select.appendChild(option);
    });

    const contributeFilter4Input = this.domElement.querySelector('#contribute-filter4-input') as HTMLSelectElement;
    contributeFilter4Input.innerHTML = '<option value="" selected style="display:none;">Content Audience</option>';
    //console.log("filter1Select", practiceInputSelect);
      this.contentAudiences.forEach(item => {
        const option = document.createElement('option');
        option.value = item;
        option.text = item;
        contributeFilter4Input.appendChild(option);
    });
  }

  private populateReleaseNumbers(selectedProduct: string): void {
    const filter5Select = this.domElement.querySelector('#filter5-select') as HTMLSelectElement;
    filter5Select.innerHTML = '<option value="" selected style="color: transparent;">Release Number</option>';

    const contributeFilter5Input = this.domElement.querySelector('#filter5-dropdown-select') as HTMLElement;
    contributeFilter5Input.innerHTML = '<label data-value style="color: transparent; display: none;" hidden>Release Number</label>';

    const uniqueReleaseNumbers = new Set<string>();

    console.log('Selected Product:', selectedProduct);
    if (selectedProduct && this.releaseNumbers[selectedProduct]) {
      this.releaseNumbers[selectedProduct].forEach(item => {
        uniqueReleaseNumbers.add(item);
      });
      console.log('Unique Release Numbers:', uniqueReleaseNumbers);
    } else {
      // Populate with all sub-practices if no practice is selected
      Object.keys(this.releaseNumbers).forEach(item => {
        this.releaseNumbers[item].forEach(item => {
          uniqueReleaseNumbers.add(item);
        });
      });
    }

    uniqueReleaseNumbers.forEach(item => {
      const option = document.createElement('option');
      option.value = item;
      option.text = item;
      filter5Select.appendChild(option);
      //filter5Selected.appendChild(option.cloneNode(true));
      const label = document.createElement('label');
      label.textContent = item;
      label.dataset.value = item;
      contributeFilter5Input.appendChild(label);
    });
  }

  private populateModules(selectedProduct: string): void {
    const filter6Select = this.domElement.querySelector('#filter6-select') as HTMLSelectElement;
    filter6Select.innerHTML = '<option value="" selected style="color: transparent;">Module</option>';
    //const filter6Selected = this.domElement.querySelector('#filter6-selected') as HTMLSelectElement;
    //filter6Selected.innerHTML = '<option value="" selected style="color: transparent;">Module</option>';

    const contributeFilter6Input = this.domElement.querySelector('#filter6-dropdown-select') as HTMLElement;
    contributeFilter6Input.innerHTML = '<label data-value style="color: transparent; display: none;" hidden>Module</label>';
  
    const uniqueModules = new Set<string>();

    console.log('Selected Product:', selectedProduct);
    if (selectedProduct && this.modules[selectedProduct]) {
      this.modules[selectedProduct].forEach(item => {
        uniqueModules.add(item);
      });
      console.log('Unique Modules:', uniqueModules);
    } else {
      // Populate with all sub-practices if no practice is selected
      Object.keys(this.modules).forEach(item => {
        this.modules[item].forEach(item => {
          uniqueModules.add(item);
        });
      });
    }

    uniqueModules.forEach(item => {
      const option = document.createElement('option');
      option.value = item;
      option.text = item;
      filter6Select.appendChild(option);
      //filter6Selected.appendChild(option.cloneNode(true));
      const label = document.createElement('label');
      label.textContent = item;
      label.dataset.value = item;
      contributeFilter6Input.appendChild(label);
    });
  }
  //---------------------------------------- END: Fetch Data and Populate the Dropdowns ----------------------------------------//

  

  //------------------------------------------ START: All Funtions Related to Filters ------------------------------------------//
  private addFilterBoxEventListeners(): void {
    const searchInput = this.domElement.querySelector('#search-input') as HTMLInputElement;
    const filterBox = this.domElement.querySelector('#filter-box') as HTMLDivElement;
  
    //When the search bar is clicked, show the filter box that contains all the filters
    searchInput.addEventListener('click', (event) => {
      event.stopPropagation(); // Prevent the click event from bubbling up to the document
      filterBox.style.display = 'block';
    });
  
    // Click anywhere else on the page and the filter box will be hidden
    document.addEventListener('click', (event) => {
      const target = event.target as Node;
      if (!searchInput.contains(target) && !filterBox.contains(target)) {
        filterBox.style.display = 'none';
      }
    });
  
    // Ensure clicks inside the filter box do not close it
    filterBox.addEventListener('click', (event) => {
      event.stopPropagation();
    });

    //Close all dropdowns if anywhere in the filter box is clicked
    if (filterBox) {
      filterBox.addEventListener('click', () => {
        //console.log('Filter box clicked');
        this.closeAllDropdowns();
      });
    }

  }

  //This sets the selected options in the filter dropdowns and displays them in a dark color
  private addSelectChangeListener(): void {
    const selectElements = this.domElement.querySelectorAll('.category-select') as NodeListOf<HTMLSelectElement>;
  
    selectElements.forEach(selectElement => {
      selectElement.addEventListener('change', () => {
        if (selectElement.value) {
          selectElement.style.color = '#222';
        } else {
          selectElement.style.color = '#999';
        }
      });
  
      // Set initial color based on the selected value
      if (selectElement.value) {
        selectElement.style.color = '#222';
      } else {
        selectElement.style.color = '#999';
      }
    });
  }

  private closeAllDropdowns(): void {
    const allDropdownContents = this.domElement.querySelectorAll('.dropdown-content');
    allDropdownContents.forEach(content => {
      (content as HTMLElement).style.display = 'none';
    });
  }

  private addProductChangeListener(): void {
    const filter2Select = this.domElement.querySelector('#filter2-select') as HTMLSelectElement;
    const contributeFilter2Input = this.domElement.querySelector('#contribute-filter2-input') as HTMLSelectElement;
    const filter5Select = this.domElement.querySelector('#filter5-select') as HTMLSelectElement;
    const productCategorySelect = this.domElement.querySelector('#filter2-select') as HTMLSelectElement;

    //Based on the selected product category, populate the options of modules, and release numbers
    filter2Select.addEventListener('change', () => {
      console.log('Product changed');
      const selectedProduct = filter2Select.value;
      this.populateModules(selectedProduct);
      this.populateReleaseNumbers(selectedProduct);
      this.removeSelectedDependantFilterDropdownItems();
    });

    contributeFilter2Input.addEventListener('change', () => {
      const selectedProduct = contributeFilter2Input.value;
      this.populateSubPracticeContributeOptions(selectedProduct);
      this.populateProductCategoryContributeOptions(selectedProduct);
      this.removeSelectedDependantFilterDropdownItems();
    });
  }

  //When the product name is changed, remove the selected options in the dependent dropdowns
  private removeSelectedDependantFilterDropdownItems(): void {
    // add the IDs of the dependent dropdowns that need to be cleared when the product name is changed
    const elementIds = [
      'filter5-dropdown-select',
      'filter6-dropdown-select'
    ];
  
    elementIds.forEach(elementId => {
      const element = this.domElement.querySelector(`#${elementId}`) as HTMLDivElement;
      if (element) {
        const nextSibling = element.nextElementSibling as HTMLDivElement;
        if (nextSibling) {
          const selectedItems = nextSibling.querySelectorAll('.selected-item');
          selectedItems.forEach(item => item.remove());
        }
      } else {
        console.error(`Element with ID "${elementId}" not found.`);
      }
    });
  }

  private addMultiSelectFeaturesToDropdowns(): void {
    const dropdowns = this.domElement.querySelectorAll('.dropdown');
    //console.log('All Dropdowns:', dropdowns);
  
    dropdowns.forEach(dropdown => {
      const selectElement = dropdown.querySelector('select') as HTMLSelectElement;
      const dropdownBtn = dropdown.querySelector('.dropdown-btn') as HTMLElement;
      const dropdownContent = dropdown.querySelector('.dropdown-content') as HTMLElement;
      const selectedItemsContainer = dropdown.querySelector('.selected-items') as HTMLElement;
  
      if (selectElement && dropdownBtn && dropdownContent && selectedItemsContainer) {
        // Populate the dropdown-content with options from the <select> element
        Array.from(selectElement.options).forEach(option => {
          const label = document.createElement('label');
          label.textContent = option.textContent;
          label.dataset.value = option.value;
          dropdownContent.appendChild(label);
        });

        // Toggle dropdown content on button click
        dropdownBtn.addEventListener('click', (event) => {
          event.stopPropagation(); // Prevent event from bubbling
          const isVisible = getComputedStyle(dropdownContent).display === 'block';
          
          // Close other dropdowns and open current dropdown
          closeAllDropdowns();
          
          if (!isVisible) {
            dropdownContent.style.display = 'block'; // Open if it was closed
          }
        });
  
        // Handle option selection
        dropdownContent.addEventListener('click', (event) => {
          event.stopPropagation(); // Prevent event bubbling
          const label = event.target as HTMLElement;
          const value = label.dataset.value;
          console.log('Selected value:', value);
  
          if (value) {
            const option = Array.from(selectElement.options).find(opt => opt.value === value);
            if (option) {
              option.selected = !option.selected;
  
              if (option.selected) {
                //addSelectedFilterTag(value);
                this.addSelectedFilterTag(value, selectedItemsContainer,selectElement);
              } else {
                //removeSelectedFilterTag(value);
                this.removeSelectedFilterTag(value,true, selectedItemsContainer,selectElement);
              }
            }
          }
        });
      }
    });
    // Close all dropdowns when clicking outside
    window.addEventListener('click', () => closeAllDropdowns());
  
    // Helper function to close all open dropdowns
    function closeAllDropdowns() {
      const allDropdownContents = document.querySelectorAll('.dropdown-content');
      allDropdownContents.forEach(content => {
        (content as HTMLElement).style.display = 'none';
      });
    }
  }

  private async addSelectedFilterTag(value: string, selectedItemsContainer: HTMLElement, selectElement: HTMLSelectElement): Promise<void> {
    console.log('Adding selected item:', value);
    console.log('Selected items container:', selectedItemsContainer);
    console.log('Select element:', selectElement);
    
    const spans = selectedItemsContainer.querySelectorAll('span');
    let matchingSpan: HTMLSpanElement | null = null;

    spans.forEach(span => {
      if (span.textContent === value) {
        matchingSpan = span;
      }
    });

    if (matchingSpan) {
      console.log('Found matching span:', matchingSpan);
    } else {
      console.log('No matching span found for value:', value);
      const item = document.createElement('div');
      item.classList.add('selected-item', `${styles.selected_item}`);
      item.innerHTML = `
        <span>${value}</span>
        <button class="remove-item ${styles.remove_item}" data-value="${value}">×</button>
      `;
      selectedItemsContainer.appendChild(item);

      // Add event listener to the close button
      item.querySelector('.remove-item')?.addEventListener('click', (event) => {
        const valueToRemove = (event.target as HTMLElement).getAttribute('data-value');
        if (valueToRemove) {
          this.removeSelectedFilterTag(valueToRemove, true, selectedItemsContainer,selectElement);
        }
      });
    }
    selectElement.dispatchEvent(new Event('change'));  
  }

  private async removeSelectedFilterTag(value: string, unselect = false, selectedItemsContainer: HTMLElement,selectElement: HTMLSelectElement) : Promise <void> {
    const items = selectedItemsContainer.querySelectorAll('.selected-item');
    items.forEach(item => {
      if (item.querySelector('span')?.textContent === value) {
        item.remove();
      }
    });

    if (unselect) {
      const option = Array.from(selectElement.options).find(opt => opt.value === value);
      if (option) {
        option.selected = false;
      }
    }

      // Mapping between selectElement IDs and corresponding div IDs
      const divMapping: { [key: string]: string } = {
        'filter2-selected': 'filter2-dropdown-select',
        'filter3-selected': 'filter3-dropdown-select',
        'filter4-selected': 'filter4-dropdown-select',
        'filter5-selected': 'filter5-dropdown-select'
      };

      // Get the corresponding div ID
      const correspondingDivId = divMapping[selectElement.id];
      console.log('Select element ID:', selectElement.id);
      console.log('Corresponding div ID:', correspondingDivId);
      if (correspondingDivId) {
        const correspondingDiv = document.getElementById(correspondingDivId);
        console.log('Corresponding div:', correspondingDiv);
        if (correspondingDiv) {
          const matchingLabel = correspondingDiv.querySelector(`label[data-value="${value}"]`) as HTMLLabelElement;
          console.log('Matching label:', matchingLabel);
          if (matchingLabel) {
            console.log('Clicking matching label');
            matchingLabel.click();
          }
        }
      }
    
  }
  //------------------------------------------- END: All Funtions Related to Filters -------------------------------------------//


  //----------------------------------------- START: All Funtions Related to Contribute ----------------------------------------//
  private addDropDownContributeListeners(): void {
    const contributeBtn = this.domElement.querySelector('#ContributeBtn') as HTMLButtonElement;
    //const dropDownContribute = this.domElement.querySelector('#DropDownContribute') as HTMLDivElement;
    //const dropDownHelp = this.domElement.querySelector('#DropDownHelp') as HTMLDivElement;
    //const dropdownContributeContent = document.getElementById('DropdownContributeContent');
    const contributePowerAppUrl = "https://apps.powerapps.com/play/e/default-c3c3137c-945a-4929-9494-ce7ac9ce5673/a/d9b234b8-ce5d-4902-8cf4-f569c6931a20?tenantId=c3c3137c-945a-4929-9494-ce7ac9ce5673&sourcetime=1732995879880&source=portal"

    contributeBtn.addEventListener('click', (event) => {
      window.open(contributePowerAppUrl, '_blank');
    });
  //Older code where the contribute feature was part of the SPFx webpart and NOT a powerapp
  /*
    contributeBtn.addEventListener('click', (event) => {
      event.stopPropagation(); // Prevent the click event from bubbling up to the document
      if(dropdownContributeContent){
        dropdownContributeContent.style.display = 'block'; // Show the content
      }
      if (dropDownContribute.style.display === 'none' || dropDownContribute.style.display === '') {
        dropDownContribute.style.display = 'block';
        dropDownHelp.style.display = 'none';
      } else {
        dropDownContribute.style.display = 'none';
      }
    });
  
    document.addEventListener('click', (event) => {
      const target = event.target as Node;
      if (!dropDownContribute.contains(target) && !contributeBtn.contains(target)) {
        dropDownContribute.style.display = 'none';
      }
    });
  
    dropDownContribute.addEventListener('click', (event) => {
      event.stopPropagation();
    });
    */
  }

  private addInfoIconEventListeners(): void {
    const iconDropdownPairs = [
      { iconId: 'client-info-icon', dropdownId: 'client-info-dropdown' },
      { iconId: 'year-info-icon', dropdownId: 'year-info-dropdown' },
      { iconId: 'product-category-info-icon', dropdownId: 'product-category-info-dropdown' }
    ];
  
    const closeAllDropdowns = () => {
      iconDropdownPairs.forEach(pair => {
        const infoDropdown = this.domElement.querySelector(`#${pair.dropdownId}`) as HTMLElement;
        if (infoDropdown) {
          infoDropdown.style.display = 'none';
        }
      });
    };
  
    iconDropdownPairs.forEach(pair => {
      const infoIcon = this.domElement.querySelector(`#${pair.iconId}`) as HTMLElement;
      const infoDropdown = this.domElement.querySelector(`#${pair.dropdownId}`) as HTMLElement;
  
      if (infoIcon && infoDropdown) {
        infoIcon.addEventListener('click', (event) => {
          event.stopPropagation(); // Prevent event from bubbling up
          const isVisible = getComputedStyle(infoDropdown).display === 'block';
          closeAllDropdowns(); // Close all other dropdowns
          infoDropdown.style.display = isVisible ? 'none' : 'block';
        });
  
        // Prevent closing the dropdown when clicking inside it
        infoDropdown.addEventListener('click', (event) => {
          event.stopPropagation();
        });
      }
    });
  
    // Close all dropdowns when clicking outside
    window.addEventListener('click', () => {
      closeAllDropdowns();
    });
  }

  private addSupportingFileUploadListeners(): void {
    const customFileUpload1 = this.domElement.querySelector('#custom-file-upload-1') as HTMLButtonElement;
    const fileUpload1 = this.domElement.querySelector('#file-upload-1') as HTMLInputElement;
    const fileName1 = this.domElement.querySelector('#file-name-1') as HTMLDivElement;
  
    if(customFileUpload1){
      customFileUpload1.addEventListener('click', () => {
        fileUpload1.click();
      });
    }

    if(customFileUpload1){
      fileUpload1.addEventListener('change', () => {
        if (fileUpload1.files && fileUpload1.files.length > 0) {
          const fileNames = Array.from(fileUpload1.files).map(file => file.name).join(', ');
          fileName1.textContent = fileNames;
        }
      })
    };
  }

  private addMainFileUploadListeners(): void {
    const customFileUploadMain = this.domElement.querySelector('#custom-file-upload-main') as HTMLButtonElement;
    const fileUploadMain = this.domElement.querySelector('#file-upload-main') as HTMLInputElement;
    const fileNameMain = this.domElement.querySelector('#file-name-main') as HTMLDivElement;
  
    customFileUploadMain.addEventListener('click', () => {
      fileUploadMain.click();
    });
  
    fileUploadMain.addEventListener('change', () => {
      if (fileUploadMain.files && fileUploadMain.files.length > 0) {
        fileNameMain.textContent = fileUploadMain.files[0].name;
      }
    });
  }

  private addEventHandlersForValidation(): void {
    const fileUploadMain = this.domElement.querySelector('#file-upload-main') as HTMLInputElement;
    const practiceInput = this.domElement.querySelector('#contribute-filter1-input') as HTMLSelectElement;
    const subPracticeInput = this.domElement.querySelector('#contribute-filter2-input') as HTMLSelectElement;
    const productCategoryInput = this.domElement.querySelector('#contribute-filter3-input') as HTMLSelectElement;
    const yearInput = this.domElement.querySelector('#contribute-filter4-input') as HTMLSelectElement;
    //const clientNameInput = this.domElement.querySelector('#client-input-select') as HTMLSelectElement;
    const confidentialInput = this.domElement.querySelector('#confidentialInput input[name="select-circle"]:checked') as HTMLInputElement;
    const submitBtn = document.querySelector('#ContributeSubmitBtn') as HTMLButtonElement;
  
    const validateInputs = () => {
      if (
        fileUploadMain.value &&
        practiceInput.value &&
        subPracticeInput.value &&
        productCategoryInput.value &&
        yearInput.value &&
        //clientNameInput.value &&
        confidentialInput.value
      ) {
        submitBtn.disabled = false;
      } else {
        submitBtn.disabled = true;
      }
    };
  
    fileUploadMain.addEventListener('change', validateInputs);
    practiceInput.addEventListener('change', validateInputs);
    subPracticeInput.addEventListener('change', validateInputs);
    productCategoryInput.addEventListener('change', validateInputs);
    yearInput.addEventListener('change', validateInputs);
    //clientNameInput.addEventListener('change', validateInputs);
    confidentialInput.addEventListener('change', validateInputs);
  
    // Initial validation check
    validateInputs();
  }

  private addContributeSubmitButtonListener(): void {
    const contributeBtn = this.domElement.querySelector('#ContributeBtn') as HTMLButtonElement;
    const dropDownContribute = this.domElement.querySelector('#DropDownContribute') as HTMLDivElement;
    const dropDownHelp = this.domElement.querySelector('#DropDownHelp') as HTMLDivElement;
    const submitBtn = document.querySelector('#ContributeSubmitBtn') as HTMLButtonElement;
  
    submitBtn.addEventListener('click', async () => {
      //const fileNameInput = this.domElement.querySelector('#fileNameInput') as HTMLInputElement;
      const fileUploadMain = this.domElement.querySelector('#file-upload-main') as HTMLInputElement;
      const practiceInput = this.domElement.querySelector('#contribute-filter1-input') as HTMLSelectElement;
      const subPracticeInput = this.domElement.querySelector('#contribute-filter2-input') as HTMLSelectElement;
      const productCategoryInput = this.domElement.querySelector('#contribute-filter3-input') as HTMLSelectElement;
      const yearInput = this.domElement.querySelector('#contribute-filter4-input') as HTMLSelectElement;
      const clientNameInput = this.domElement.querySelector('#client-input-select') as HTMLSelectElement;
      const keywordsInput = this.domElement.querySelector('#keywordsInput') as HTMLInputElement;
      const confidentialInput = this.domElement.querySelector('#confidentialInput input[name="select-circle"]:checked') as HTMLInputElement;
      const confidentialDetailsInput = this.domElement.querySelector('#confidentialDetailsInput') as HTMLInputElement;
      const fileInput = this.domElement.querySelector('#file-upload-1') as HTMLInputElement; //This is the supporting document
           

      //console.log('Practice Input:', practiceInput.value);
      //console.log('Sub Practice Input:', subPracticeInput.value);
      //console.log('Product Category Input:', productCategoryInput.value);
      //console.log('Year Input:', yearInput.value);
      console.log('File Upload Main:', fileUploadMain[0] ? fileUploadMain.files : 'No file selected');

      //if (!practiceInput || !subPracticeInput || !productCategoryInput || !yearInput || !fileInput || !fileInput.files || !fileInput.files.length || !fileUploadMain.files) 
      if (!practiceInput || (practiceInput.value !== "Marketing" && !subPracticeInput) || !productCategoryInput || !yearInput || !fileUploadMain.files) {
        alert('Please fill in all required fields.');
        return;
      }
  
      if (!confidentialInput) {
        alert('Please select a confidentiality option.');
        return;
      }

      /*Upload the main document. Overview: 
      1. Generate a digest so that authentication is taken care of
      2. If the file is confidential, the file is uploaded to a confidential document library. We take the file's item ID and store it in a variable (we use this item ID later in the approval workflow)
      3. A list item is created in the UserContributionFiles list. This is for KM review 
      4. If the file is not confidential, it is uploaded as an attachment to the list item.
      5. Then, upload the supporting documents to the list for KM review*/
      let parentFileItemId='';
      if(fileUploadMain && fileUploadMain.files && fileUploadMain.files.length && practiceInput && subPracticeInput && productCategoryInput && yearInput &&confidentialInput){
        const fileName = fileUploadMain.files[0].name;
        const practice = practiceInput.value;
        const subPractice = subPracticeInput.value;
        const productCategory = productCategoryInput.value;
        const year = yearInput.value;
        const clientName = clientNameInput ? clientNameInput.value : '';
        const keywords = keywordsInput ? keywordsInput.value : '';
        const confidential = confidentialInput.value === 'yes';
        const confidentialDetails = confidentialDetailsInput ? confidentialDetailsInput.value : '';

        const file = fileUploadMain.files[0];
        const fileBuffer = await file.arrayBuffer();

        console.log('File:', file);
        console.log('File Buffer:', fileBuffer);
        console.log('Confidential Input:', confidentialInput.value);
        console.log('Confidential:', confidential);
    
        try {
          // Generate a new digest
          const digestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/contextinfo`;
          const digestResponse = await fetch(digestUrl, {
            method: 'POST',
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-Type': 'application/json;odata=nometadata'
            }
          });
    
          if (!digestResponse.ok) {
            throw new Error('Error generating digest');
          }
    
          const digestData = await digestResponse.json();
          const digest = digestData.FormDigestValue; // Get the fresh digest token

          // Upload the file in a confidential document library if the file is marked as confidential
          let itemServerRelativeUrl = '';
          if(confidentialInput.value ==='yes'){
            const uploadUrl = `https://zinnovconsulting.sharepoint.com/sites/KnowledgeExchangePortal/_api/web/GetFolderByServerRelativeUrl('/sites/KnowledgeExchangePortal/TemporaryStorage')/Files/add(url='${fileName}', overwrite=true)`;
            const uploadResponse = await fetch(uploadUrl, {
            method: 'POST',
            headers: {
              'Accept': 'application/json;odata=verbose',
              'X-RequestDigest': digest,
              'Content-Length': fileBuffer.byteLength.toString()
            },
            body: fileBuffer
            });
            console.log('Raw Upload Response:', uploadResponse);
            console.log('Main File Upload Response:', JSON.stringify(uploadResponse, null, 2));            
            
            if (!uploadResponse.ok) {
              const errorText = await uploadResponse.text();
              console.error('Error uploading confidential file:', errorText);
              throw new Error('Error uploading confidential file');
            }

            const uploadData = await uploadResponse.json();
            console.log('Parsed Upload Data:', JSON.stringify(uploadData, null, 2));
            itemServerRelativeUrl = uploadData.d.ServerRelativeUrl; // Get the item ID from the response
            console.log('Uploaded file relatiive id:', itemServerRelativeUrl);
          }

          const mainFileData = {
            __metadata: { type: 'SP.Data.UserContributionFilesListItem' }, // Correct type name
            Title: fileName,
            IsSupportingDocument: 'No',
            Practice: practice,
            SubPractice: subPractice,
            ProductCategory: productCategory,
            Year: year,
            ClientName: clientName,
            Keywords: keywords,
            Confidential: confidential === true ? 'Yes' : 'No', // Ensure this is a boolean value
            ConfidentialDetails: confidentialDetails,
            Status: 'Pending', // Ensure this matches one of the choices in the Status column
            ServerRelativeUrl: itemServerRelativeUrl ? itemServerRelativeUrl : ''
          };
      
          console.log('Main File Data:', JSON.stringify(mainFileData, null, 2));
    
          // Create the main file list item
          const response = await fetch(`https://zinnovconsulting.sharepoint.com/sites/KnowledgeExchangePortal/_api/web/lists/getbytitle('UserContributionFiles')/items`, {
            method: 'POST',
            headers: {
              'Accept': 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose',
              'X-RequestDigest': digest
            },
            body: JSON.stringify(mainFileData)
          });
    
          if (!response.ok) {
            const errorText = await response.text();
            console.error('Error creating list item:', errorText);
            throw new Error('Error creating list item');
          }
    
          const item = await response.json();
          parentFileItemId = item.d.Id;
          console.log('confidentialDetailsInput.value: ', confidentialDetailsInput.value, 'parentFileItemId: ', parentFileItemId);
          
          // Upload the file as an attachment if the file is not confidential
          if(confidentialDetailsInput.value != 'yes') {
            console.log('Uploading main file as an attachment');
            const fileResponse = await fetch(`https://zinnovconsulting.sharepoint.com/sites/KnowledgeExchangePortal/_api/web/lists/getbytitle('UserContributionFiles')/items(${item.d.Id})/AttachmentFiles/add(FileName='${file.name}')`, {
              method: 'POST',
              headers: {
                'Accept': 'application/json;odata=verbose',
                'X-RequestDigest': digest
              },
              body: fileBuffer
            });
    
            if (!fileResponse.ok) {
              const errorText = await fileResponse.text();
              console.error('Error uploading file:', errorText);
              throw new Error('Error uploading file');
            }
          }
          
          // Show the thank you page with updated content
          if (fileInput.files && fileInput.files.length === 0) {
            const thankYouContentContribute = this.domElement.querySelector('#ThankYouContentContribute') as HTMLDivElement;
            const dropDownContribute = document.getElementById('DropDownContribute');
            const dropdownContributeContent = document.getElementById('DropdownContributeContent');          
            if (dropdownContributeContent) {
              dropdownContributeContent.style.display = 'none';
            } else {
              console.error('DropdownContributeContent element not found');
            }
          
            if (thankYouContentContribute) {
              thankYouContentContribute.style.display = 'block';
              setTimeout(() => {
                if (dropdownContributeContent && dropDownContribute) {
                  dropdownContributeContent.style.display = 'none';
                  dropDownContribute.style.display = 'none';
                }
                thankYouContentContribute.style.display = 'none';
              }, 3000);
            } else {
              console.error('ThankYouContentContribute element not found');
            }
          
            // Reset the values
            fileUploadMain.value = '';
            practiceInput.value = '';
            subPracticeInput.value = '';
            productCategoryInput.value = '';
            yearInput.value = '';
            clientNameInput.value = '';
            const fileNameMain = this.domElement.querySelector('#file-name-main') as HTMLDivElement;
            fileNameMain.textContent = '';
          }

        } catch (error) {
          console.error(error);
          alert('An error occurred while submitting your contribution.');
        }
      }
      
      //If there are any supporting documents, upload them to the list
      console.log('File Input: ', fileInput);
      console.log('File Input Files: ', fileInput.files);
      console.log('File Input Files Length: ', fileInput.files ? fileInput.files.length : 'no file input length');
      console.log(parentFileItemId);

      if (fileInput && fileInput.files && fileInput.files.length && fileUploadMain && practiceInput && subPracticeInput && productCategoryInput && yearInput && confidentialInput) {
        const practice = practiceInput.value;
        const subPractice = subPracticeInput.value;
        const productCategory = productCategoryInput.value;
        const year = yearInput.value;
        const clientName = clientNameInput ? clientNameInput.value : '';
        const keywords = keywordsInput ? keywordsInput.value : '';

          for (const file of Array.from(fileInput.files)) {
          const fileName = file.name;
          const fileBuffer = await file.arrayBuffer();

          const itemData = {
              __metadata: { type: 'SP.Data.UserContributionFilesListItem' }, // Correct type name
              Title: fileName,
              IsSupportingDocument: 'Yes',
              ParentFileItemID: parentFileItemId.toString(),
              Practice: practice,
              SubPractice: subPractice,
              ProductCategory: productCategory,
              Year: year,
              ClientName: clientName,
              Keywords: keywords,
              Confidential: 'No', // Ensure this is a boolean value
              Status: 'Pending' // Ensure this matches one of the choices in the Status column
          };

          console.log('Item Data:', itemData);

          try {
              // Generate a new digest
              const digestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/contextinfo`;
              const digestResponse = await fetch(digestUrl, {
                  method: 'POST',
                  headers: {
                      'Accept': 'application/json;odata=nometadata',
                      'Content-Type': 'application/json;odata=nometadata'
                  }
              });

              if (!digestResponse.ok) {
                  throw new Error('Error generating digest');
              }

              const digestData = await digestResponse.json();
              const digest = digestData.FormDigestValue; // Get the fresh digest token

              // Create the list item
              const response = await fetch(`https://zinnovconsulting.sharepoint.com/sites/KnowledgeExchangePortal/_api/web/lists/getbytitle('UserContributionFiles')/items`, {
                  method: 'POST',
                  headers: {
                      'Accept': 'application/json;odata=verbose',
                      'Content-Type': 'application/json;odata=verbose',
                      'X-RequestDigest': digest
                  },
                  body: JSON.stringify(itemData)
              });

              if (!response.ok) {
                  const errorText = await response.text();
                  console.error('Error creating list item:', errorText);
                  throw new Error('Error creating list item');
              }

              const item = await response.json();

              // Upload the file as an attachment
              const fileResponse = await fetch(`https://zinnovconsulting.sharepoint.com/sites/KnowledgeExchangePortal/_api/web/lists/getbytitle('UserContributionFiles')/items(${item.d.Id})/AttachmentFiles/add(FileName='${file.name}')`, {
                  method: 'POST',
                  headers: {
                      'Accept': 'application/json;odata=verbose',
                      'X-RequestDigest': digest
                  },
                  body: fileBuffer
              });

              if (!fileResponse.ok) {
                  const errorText = await fileResponse.text();
                  console.error('Error uploading file:', errorText);
                  throw new Error('Error uploading file');
              }
          } catch (error) {
              console.error(error);
              alert('An error occurred while submitting your contribution.');
          }
        }
      }
      const thankYouContentContribute = this.domElement.querySelector('#ThankYouContentContribute') as HTMLDivElement;
      const dropDownContribute = document.getElementById('DropDownContribute');
      const dropdownContributeContent = document.getElementById('DropdownContributeContent');
      const fileNameMain = this.domElement.querySelector('#file-name-main') as HTMLDivElement;
    
      if (dropdownContributeContent) {
        dropdownContributeContent.style.display = 'none';
      } else {
        console.error('DropdownContributeContent element not found');
      }
    
      if (thankYouContentContribute) {
        thankYouContentContribute.style.display = 'block';
        setTimeout(() => {
          if (dropdownContributeContent && dropDownContribute) {
            dropdownContributeContent.style.display = 'none';
            dropDownContribute.style.display = 'none';
          }
          thankYouContentContribute.style.display = 'none';
        }, 3000);

      //reset the values
      fileUploadMain.value = '';
      practiceInput.value = '';
      subPracticeInput.value = '';
      productCategoryInput.value = '';
      yearInput.value = '';
      clientNameInput.value = '';
      keywordsInput.value = '';
      confidentialInput.value = '';
      confidentialDetailsInput.value = '';
      fileInput.value = '';
      fileInput.textContent = '';
      fileNameMain.textContent = '';
      }
    });
  }
  //----------------------------------------- END: All Funtions Related to Contribute ------------------------------------------//



  //-------------------------------------------- START: All Funtions Related to Help -------------------------------------------//
  private DropHelpButtonListener(): void {
    const helpBtn = this.domElement.querySelector('#HelpBtn') as HTMLButtonElement;
    const dropDownHelp = this.domElement.querySelector('#DropDownHelp') as HTMLDivElement;
    const dropDownContribute = this.domElement.querySelector('#DropDownContribute') as HTMLDivElement;
    const submitBtn = this.domElement.querySelector('#SubmitBtn') as HTMLButtonElement;
    const helpContent = this.domElement.querySelector('#HelpContent') as HTMLDivElement;
    const thankYouContent = this.domElement.querySelector('#ThankYouContent') as HTMLDivElement;
  
    helpBtn.addEventListener('click', (event) => {
      //This is the see and ensure that the submit button is enabled only if the category is selected and a query is typed in
      const categorySelect = document.querySelector('.category-select') as HTMLSelectElement;
      const commentsTextarea = document.querySelector('#queryTextarea') as HTMLTextAreaElement;

      //console.log('Selected Help Category: ' ,categorySelect.value)
      //console.log('Comments in help: ' ,commentsTextarea.value.trim())

      const toggleSubmitButton = () => {
        if (!categorySelect.value || !commentsTextarea.value.trim()) {
          submitBtn.disabled = true;
        } else {
          submitBtn.disabled = false;
        }
      };

      // Initial check
      toggleSubmitButton();

      // Add event listeners
      categorySelect.addEventListener('change', toggleSubmitButton);
      commentsTextarea.addEventListener('input', toggleSubmitButton);
      //end of the check

      event.stopPropagation(); // Prevent the click event from bubbling up to the document
      if (dropDownHelp.style.display === 'none' || dropDownHelp.style.display === '') {
        dropDownHelp.style.display = 'block';
        if (dropDownContribute) {
          dropDownContribute.style.display = 'none';
        }
      } else {
        dropDownHelp.style.display = 'none';
      }
    });
  
    document.addEventListener('click', (event) => {
      const target = event.target as Node;
      if (!dropDownHelp.contains(target) && !helpBtn.contains(target)) {
        dropDownHelp.style.display = 'none';
      }
    });
  
    dropDownHelp.addEventListener('click', (event) => {
      event.stopPropagation();
    });
  
    submitBtn.addEventListener('click', async (event) => {
      event.preventDefault();
    
      // Send mail to the KM team based on the selections and the input text
      const categorySelect = document.querySelector('.category-select') as HTMLSelectElement;
      const commentsTextarea = document.querySelector('#queryTextarea') as HTMLTextAreaElement;
      const dropDownHelp = this.domElement.querySelector('#DropDownHelp') as HTMLDivElement;
    
      if (categorySelect && commentsTextarea) {
        console.log('entered the submit button click event');
        const selectedCategory = categorySelect.value;
        const comments = commentsTextarea.value;
        
        //This function calls the ms graph client to send the mail to the knowledge mail id from the maillbox of the logged in user
        await this.initateHelpQueryProcess(selectedCategory, comments);

        // Show the thank you page with updated content
        const thankYouP = thankYouContent.querySelector('p') as HTMLParagraphElement;
        const thankYouH4 = thankYouContent.querySelector('h4') as HTMLHeadingElement;
        thankYouP.innerText = "We will get back to you shortly.";
        thankYouH4.innerText = "Thank you for your query!";

        helpContent.style.display = 'none';
        thankYouContent.style.display = 'block';
    
        // Reset the thank you page to the help page after 5 seconds
        setTimeout(() => {
          helpContent.style.display = 'block';
          thankYouContent.style.display = 'none';
          dropDownHelp.style.display = 'none';

          // Reset the content back to the original message
          const thankYouP = thankYouContent.querySelector('p') as HTMLParagraphElement;
          const thankYouH4 = thankYouContent.querySelector('h4') as HTMLHeadingElement;
      
          thankYouP.innerText = "We are in the process of verifying the document, and will update it on the portal soon.";
          thankYouH4.innerText = "Thank you for your contribution"; 

          // Reset the values of category select and query text area
          categorySelect.value = "";
          commentsTextarea.value = "";    
        }, 3000);// Adjust the delay as needed
      } else {
        console.error('Category select or comments textarea not found.');
      }
    });
  }

  //Function to send email using MS Graph for the help functionality
  private async initateHelpQueryProcess(subject: string, body: string): Promise<void> {
    // Create an item in the tracker list
    const currentTime = new Date();
    const offset = 5.5 * 60; // GMT+5:30 in minutes
    const currentTimeInGMT530 = new Date(currentTime.getTime() + offset * 60 * 1000).toISOString();

    // Fetch the request digest
    const digestResponse = await fetch(`${this.context.pageContext.web.absoluteUrl}/_api/contextinfo`, {
        method: 'POST',
        headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose'
        }
    });

    if (!digestResponse.ok) {
        throw new Error(`HTTP error! status: ${digestResponse.status}`);
    }

    const digestData = await digestResponse.json();
    const requestDigest = digestData.d.GetContextWebInformation.FormDigestValue;

    // Send a post request to the list
    const listApiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Help_UserQueries')/items`;
    const listData = {
        __metadata: { type: 'SP.Data.Help_x005f_UserQueriesListItem' },
        Title: this.context.pageContext.user.email,
        QueryRaisedOn: currentTimeInGMT530,
        UserName: this.context.pageContext.user.displayName,
        Subject: subject,
        QueryBody: body
    };

    const response = await fetch(listApiUrl, {
        method: 'POST',
        headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'X-RequestDigest': requestDigest
        },
        body: JSON.stringify(listData)
    });

    if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
    }

    console.log('Help Query Submitted. Response: ', response);

    const linkToQueryItemInList = response.headers.get('Location');
    console.log('Link to query item in list: ', linkToQueryItemInList);
  }
  //--------------------------------------------- END: All Funtions Related to Help --------------------------------------------//
  


  //------------------------------------------ START: All Funtions Related to Download -----------------------------------------//
  /*
  private async handleDownloadButtons(): Promise<void> {
    //const selectedDownloadItems: { title: string; linkUrl: string }[] = [];
    const addDownloadCheckBoxListener = async () => {
      //console.log('Adding event listeners to checkboxes');
      // Select all checkboxes
      const checkboxes = document.querySelectorAll(`.${styles.Checkbox}`);
      checkboxes.forEach(async (checkbox) => {
        checkbox.addEventListener('change', (event) => {
          const target = event.target as HTMLInputElement;
          const title = target.getAttribute('data-title');
          let linkUrl = target.getAttribute('data-linkurl');
          let knowledgeCategory = target.getAttribute('data-knowledgecategory');
    
          if (title && linkUrl && knowledgeCategory) {
            if (target.checked) {
              // Remove ?web=1 if it exists at the end of the URL
              if (linkUrl.endsWith('?web=1')) {
                linkUrl = linkUrl.replace(/\?web=1$/, '');
              }
              this.selectedDownloadItems.push({ title, linkUrl , knowledgeCategory});
              //console.log('Selected items:', selectedDownloadItems);
            } else {
              // Remove ?web=1 if it exists at the end of the URL
              if (linkUrl.endsWith('?web=1')) {
                linkUrl = linkUrl.replace(/\?web=1$/, '');
              }
              const index = this.selectedDownloadItems.findIndex(item => item.title === title && item.linkUrl === linkUrl);
              if (index !== -1) {
                this.selectedDownloadItems.splice(index, 1);
              }
            }
           //console.log(selectedDownloadItems); // For debugging purposes
          }
        });
      });
    };

    const downloadAsCSV = async () => {
      if (this.isDownloading) {
        return; // If a download is already in progress, do nothing
      }
  
      if (this.selectedDownloadItems.length === 0) {
        alert('No items selected');
        return;
      }

      await this.updateDownloadAnalytics();
  
      this.isDownloading = true; // Set the flag to indicate that the download is in progress
  
      const csvContent = 'data:text/csv;charset=utf-8,' + this.selectedDownloadItems.map(item => `${item.title},${item.linkUrl}, ${item.knowledgeCategory}`).join('\n');
      const encodedUri = encodeURI(csvContent);
      const link = document.createElement('a');
      link.setAttribute('href', encodedUri);
      link.setAttribute('download', 'searchResults.csv');
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
  
      this.isDownloading = false; // Reset the flag after the download completes
    };
    
    const downloadAsZIP = async () => {
      if (this.isDownloading) {
        return; // If a download is already in progress, do nothing
      }
  
      if (this.selectedDownloadItems.length === 0) {
        alert('No items selected');
        return;
      }

      await this.updateDownloadAnalytics();
  
      this.isDownloading = true; // Set the flag to indicate that the download is in progress
  
      // Example ZIP download logic
      const zip = new JSZip();
      this.selectedDownloadItems.forEach(item => {
        zip.file(item.title, fetch(item.linkUrl).then(res => res.blob()));
      });
  
      zip.generateAsync({ type: 'blob' }).then(content => {
        const link = document.createElement('a');
        link.href = URL.createObjectURL(content);
        link.download = 'searchResults.zip';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
  
        this.isDownloading = false; // Reset the flag after the download completes
      }).catch(error => {
        console.error('Error generating ZIP:', error);
        this.isDownloading = false; // Reset the flag in case of error
      });
    };

    addDownloadCheckBoxListener();
    document.getElementById('downloadAsCsvBtn')?.addEventListener('click', downloadAsCSV);
    document.getElementById('downloadAsZipBtn')?.addEventListener('click', downloadAsZIP);
  }

  private removeAllDownloadListenersById(id: string, event: string) {
    const elements = document.querySelectorAll(`#${id}`);
    elements.forEach(element => {
      const newElement = element.cloneNode(true);
      element.parentNode?.replaceChild(newElement, element);
    });
  }

  private async updateDownloadAnalytics(): Promise<void> {
    const pathToAnalyticsCsvFile = "https://zeomega.sharepoint.com/sites/zehub/Analytics/Analytics_FileDownloads.csv";

    try {
      for (const item of this.selectedDownloadItems) {
        const userEmail = this.context.pageContext.user.email;
        const timeStamp = new Date();
        const offset = 5.5 * 60; // GMT+5:30 in minutes
        let gmt530TimeStamp = new Date(timeStamp.getTime() + offset * 60 * 1000);
    
        const day = gmt530TimeStamp.getUTCDate().toString().padStart(2, '0');
        const month = (gmt530TimeStamp.getUTCMonth() + 1).toString().padStart(2, '0'); // Months are zero-based
        const yearr = gmt530TimeStamp.getUTCFullYear();
        const hours = gmt530TimeStamp.getUTCHours().toString().padStart(2, '0');
        const minutes = gmt530TimeStamp.getUTCMinutes().toString().padStart(2, '0');
        const seconds = gmt530TimeStamp.getUTCSeconds().toString().padStart(2, '0');
        const milliseconds = gmt530TimeStamp.getUTCMilliseconds().toString().padStart(3, '0');
    
        const formattedDate = `${yearr}-${month}-${day}T${hours}:${minutes}:${seconds}.${milliseconds}Z`;
    
        console.log(formattedDate);

        // Fetch the existing CSV file
        const response = await fetch(pathToAnalyticsCsvFile);
        if (!response.ok) {
          throw new Error('Failed to fetch the CSV file');
        }
        const csvText = await response.text();

        // Parse the CSV data
        const rows = csvText.split('\n');
        const newRow = [
          userEmail,
          formattedDate,
          item.title,
          item.linkUrl,
          item.knowledgeCategory
        ];
        const updatedRows = [...rows, newRow.join(',')];

        // Convert the updated data back to CSV format
        const updatedCsvText = updatedRows.join('\n');

        // Write the updated CSV data back to the file
        const writeResponse = await fetch(pathToAnalyticsCsvFile, {
          method: 'PUT',
          headers: {
            'Content-Type': 'text/csv'
          },
          body: updatedCsvText
        });

        if (!writeResponse.ok) {
          throw new Error('Failed to write to the CSV file');
        }
        console.log('Filedownloads Analytics file updated successfully');
      }
      console.log('New row added to the CSV file');
    } catch (error) {
      console.error('Error adding row to CSV file:', error);
    }
  }
  */
  //------------------------------------------- END: All Funtions Related to Download ------------------------------------------//



  //------------------------------------------- START: All Funtions Related to Search ------------------------------------------//
  private async identifySearchTagsForRendering(filters: { [key: string]: string | string[] }): Promise<void> {
    //This method identifies the search tags that will be rendered in the search container based on the selected. Then it calls the createSearchResultTags method to render the tags
    const searchTagsContainer = this.domElement.querySelector('#search-tags') as HTMLDivElement;
  
    if (searchTagsContainer) {
      searchTagsContainer.innerHTML = ''; // Clear existing tags
      console.log('Filters:', filters);
  
      Object.keys(filters).forEach(key => {
        console.log('Key:', key);
        const values = filters[key];
        console.log('Values:', values);
  
        if (values) {
          if (typeof values === 'string') {
            // Handle single string value
            this.createSearchResultTags(searchTagsContainer, key, values);
          } else if (Array.isArray(values)) {
            // Handle array of string values
            values.forEach(value => {
              this.createSearchResultTags(searchTagsContainer, key, value);
            });
          }
        }
      });
    } else {
      console.error('Element with ID "search-tags" not found.');
    }
  }
  
  private createSearchResultTags(container: HTMLDivElement, key: string, value: string): void {
    //This method creates the necessary filter tags in the search container based on the selected filters
    if (value && value !== 'Departments' && value !== 'Product Name' && value !== 'Knowledge Category' && value !== 'Content Audience' && value !== 'Release Number' && value !== 'Module' && value !== 'Additional Keywords') {
      
      //Each keyword is split by space or comma and a tag is created for each keyword
      if (key === 'keywordInput') {
        // Split the keywords by space or comma
        const keywords = value.split(/[\s,]+/).filter(keyword => keyword.trim() !== '');
        keywords.forEach(keyword => {
          const span = document.createElement('span');
          span.innerHTML = `${keyword} <i class="fa fa-times close-icon"></i>`;
          container.appendChild(span);
          // Add event listener to the <i> tag. This will remove the tag when clicked and trigger a new search
          const closeIcon = span.querySelector('i');
          if (closeIcon) {
            this.closeSearchResultTagsListener(closeIcon, key, keyword);
          }
        });
      } else {
        const span = document.createElement('span');
        span.innerHTML = `${value} <i class="fa fa-times close-icon"></i>`;
        container.appendChild(span);
        // Add event listener to the <i> tag. This will remove the tag when clicked and trigger a new search
        const closeIcon = span.querySelector('i');
        if (closeIcon) {
          this.closeSearchResultTagsListener(closeIcon, key, value);
        }
      }
    }
  }

  private closeSearchResultTagsListener(closeIcon: HTMLElement, key: string, value: string): void {
    // Mapping of key values to select element IDs
   const keyToIdMap: { [key: string]: string } = {
     filter1Select: 'filter1-select',
     filter2Select: 'filter2-select',
     filter3Select: 'filter3-select',
     filter4Select: 'filter4-select',
     filter5Select: 'filter5-select',
     filter6Select: 'filter6-select',
     keywordInput: 'keyword-input'
   };
   
   closeIcon.addEventListener('click', async () => {
     const selectElementId = keyToIdMap[key];
     const getSelectedValues = (selectElement: HTMLSelectElement): string[] => {
       return Array.from(selectElement.selectedOptions).map(option => option.value);
     };
     if (selectElementId) {
       console.log('entered close icon')
       const selectElement = this.domElement.querySelector(`#${selectElementId}`) as HTMLSelectElement;
       console.log('Option to Deselect: ', value);
       //console.log('Select Element: ', selectElement);


       if (selectElement) {
        console.log('Select element found');
        if (key === 'keywordInput') {
          
          await new Promise<void>(async (resolve) => {
          const keywords = selectElement.value.split(/[\s,]+/).filter(keyword => keyword.trim() !== '');
          //console.log('Keywords:', keywords);
          //console.log('Value:', value);
          const updatedKeywords = keywords.filter(keyword => keyword !== value);
          //console.log('Updated keywords:', updatedKeywords);
          selectElement.value = updatedKeywords.join(' '); ; // Clear the input field
          //console.log('Select element value:', selectElement.value);  
          resolve();
          });
        } else {
          const selectElementOptions = getSelectedValues(this.domElement.querySelector(`#${selectElementId}`) as HTMLSelectElement);
          console.log('Select Element id: ', selectElementId);
          console.log('Select Element Options: ', selectElementOptions);
          const selectedItems = this.domElement.querySelectorAll('.selected-item');
          console.log('Selected Items:', selectedItems);
        
          if (selectElementOptions.length >= 1) {
            // Handle multiple selected options
            const options = Array.from(selectElement.options);
            //console.log('Options:', options);
            options.forEach(option => {
              if (option.value === value) {
                console.log('Deselecting option:', option);
                option.selected = false;
              }
            });
            selectedItems.forEach(item => {
              const span = item.querySelector('span');
              if (span && span.textContent === value) {
                item.remove();
              }
          });
          } /*else {
          //console.log('Deselecting for single selected option:', value);
          selectElement.selectedIndex = 0; // Reset to the first option
          }
          */
        }
       await this.handleSearchResults();
       } else {
         console.error(`Select element with ID "${selectElementId}" not found.`);
       }
     } else {
       console.error(`No mapping found for key "${key}".`);
     }

      // Remove any <span> tag inside the specific <div> that has the value matching the value parameter
     const searchTagsContainer = this.domElement.querySelector('#search-tags') as HTMLDivElement;
     if (searchTagsContainer) {
       const spans = searchTagsContainer.querySelectorAll('span');
       spans.forEach(span => {
         if (span.textContent && span.textContent.includes(value)) {
           searchTagsContainer.removeChild(span);
           console.log('Removed span:', span);
         }
       });
     } else {
       console.error('Element with ID "search-tags" not found.');
     }
   });
 }

  private async handleSearchResults(): Promise<void> {
    this.showLoader();
    const searchInput = (this.domElement.querySelector('#search-input') as HTMLInputElement).value;
    this.searchTerm = searchInput;
    this.searchResultsNumber = 0
    const searchResultsHeading = document.getElementById('search-results-heading');
      if(searchResultsHeading){
        searchResultsHeading.innerHTML = `${this.homeIconHtml} Search Results for '${this.searchTerm}' (${this.searchResultsNumber})`;
      }
    
    const searchResultsContainer = this.domElement.querySelector('#search-result') as HTMLDivElement;
    searchResultsContainer.innerHTML = ''

    const filterBox = this.domElement.querySelector('#filter-box') as HTMLDivElement;
    filterBox.style.display = 'none';
    
    const searchBtn = this.domElement.querySelector('#search-btn') as HTMLDivElement;
    const searchResults = this.domElement.querySelector('#search-results') as HTMLDivElement;
    //const searchResultBannerFilter = this.domElement.querySelector('#search-result-banner-filter') as HTMLDivElement;

    const getSelectedValues = (selectElement: HTMLSelectElement): string[] => {
      return Array.from(selectElement.selectedOptions).map(option => option.value);
    };
    
    // Retrieve selected filters
    const keywordInput = (this.domElement.querySelector('#keyword-input') as HTMLInputElement).value;
    
    const filter1Select = getSelectedValues(this.domElement.querySelector('#filter1-select') as HTMLSelectElement);
    const filter2Select = getSelectedValues(this.domElement.querySelector('#filter2-select') as HTMLSelectElement);
    const filter3Select = getSelectedValues(this.domElement.querySelector('#filter3-select') as HTMLSelectElement);
    const filter4Select = getSelectedValues(this.domElement.querySelector('#filter4-select') as HTMLSelectElement);
    const filter5Select = getSelectedValues(this.domElement.querySelector('#filter5-select') as HTMLSelectElement);
    const filter6Select = getSelectedValues(this.domElement.querySelector('#filter6-select') as HTMLSelectElement);
    
    console.log('Selected Department:', filter1Select);
    console.log('Selected Product Names:', filter2Select); 
    console.log('Selected Knowledge Category:', filter3Select);
    console.log('Selected Content Audience:', filter4Select);
    console.log('Selected Release Numbers:', filter5Select);
    console.log('Selected Module:', filter6Select);
    console.log('Keyword:', keywordInput);

      await new Promise<void>(async (resolve) => {
        console.log("filter1Select === '': ", filter1Select.length === 0); 
        console.log("filter5SelectOptions: ", filter5Select);
        console.log("filter5SelectOptions.length: ", filter5Select.length);
        console.log("filter5SelectOptions.length===0: ", filter5Select.length===0);
        console.log("productCategorySelectOptions.length===0: ", filter3Select.length===0);
        console.log("filter4SelectOptions.length===0: ", filter4Select.length===0);
        console.log("filter5SelectOptions.length===0: ", filter5Select.length===0);
        console.log("keywordInput === '': ", keywordInput === '');
    //if(filter1Select === '' && filter5Select === '' && productCategorySelect === '' && filter4Select === '' && filter5Select === '' && keywordInput === '') {
    if(searchInput !== ''){
      await this.setSelectedOptions(filter1Select, filter2Select, filter3Select, filter4Select, filter5Select, filter6Select);
      //searchResultBannerFilter.style.display = 'flex';
      searchResults.style.display = 'flex'; 
      if(filter1Select.length<=1 && filter2Select.length<=1 && filter5Select.length<=1 && filter3Select.length<=1 && filter4Select.length<=1 && filter6Select.length<=1 && keywordInput === '') {
        console.log('All filters are empty but search input is not empty');
        //searchResultBannerFilter.style.display = 'none';
        //searchResults.style.display = 'none'; 
        await this.identifySearchTagsForRendering({ filter1Select,filter2Select, filter3Select, filter4Select,filter5Select,filter6Select, keywordInput });
        await this.searchSharePoint();
      }else{
        console.log('Practice is not marketing');
        await this.identifySearchTagsForRendering({  filter1Select,filter2Select, filter3Select, filter4Select,filter5Select,filter6Select, keywordInput });
        this.searchResultsNumber = 0
          if (this.searchResultsNumber){
            const searchResultsHeading = document.getElementById('search-results-heading');
            if(searchResultsHeading){
              searchResultsHeading.innerHTML = `${this.homeIconHtml} Search Results for '${this.searchTerm}' (${this.searchResultsNumber})`;
            }
          }
        await this.searchSharePoint();
        const heading = document.getElementById('search-results-heading');
          if (heading) {
            heading.textContent = `${this.homeIconHtml} Search Results for '${this.searchTerm}' (${this.searchResultsNumber})`;
          };
      }
    }else{
      console.log('Search input is empty');
      searchResults.style.display = 'none'; 
    }
      resolve();
    });
      this.addLoadMoreFunctionality();
      this.addHomeIconEventHandler();
      this.hideLoader();
      //this.handleDownloadButtons();
      this.jumpToSearchResults();
      this.logSearchDataForAnalytics();
  }

  //Function to log the search data information to a file for analytics
  private logSearchDataForAnalytics = async (): Promise<void> => {
    console.log('Logging search data...')
    const pathToAnalyticsCsvFile = "https://zeomega.sharepoint.com/sites/zehub/Analytics/Analytics_SearchData.csv"

    const getSelectedValues = (selectElement: HTMLSelectElement): string[] => {
      return Array.from(selectElement.selectedOptions).map(option => option.value);
    };

    // Retrieve selected filters
    const searchInput = (this.domElement.querySelector('#search-input') as HTMLInputElement).value.replace(/,/g, ';');
    const filter1Select = getSelectedValues(this.domElement.querySelector('#filter1-select') as HTMLSelectElement).join(';');
    const filter2Select = getSelectedValues(this.domElement.querySelector('#filter2-select') as HTMLSelectElement).join(';');
    const filter3Select = getSelectedValues(this.domElement.querySelector('#filter3-select') as HTMLSelectElement).join(';');
    const filter4Select = getSelectedValues(this.domElement.querySelector('#filter4-select') as HTMLSelectElement).join(';');
    const filter5Select = getSelectedValues(this.domElement.querySelector('#filter5-select') as HTMLSelectElement).join(';');
    const filter6Select = getSelectedValues(this.domElement.querySelector('#filter6-select') as HTMLSelectElement).join(';');
    const keywordInput = (this.domElement.querySelector('#keyword-input') as HTMLInputElement).value.replace(/,/g, ';');

    const searchResultsNumber = this.searchResultsNumber;
    const userEmail = this.context.pageContext.user.email;
    const timeStamp = new Date().toLocaleString().replace(/,/g, ';');

    /*
    const searchInput = (this.domElement.querySelector('#search-input') as HTMLInputElement).value.replace(/,/g, ';');
    const practiceSelect = (this.domElement.querySelector('#practice-select') as HTMLSelectElement).value.replace(/,/g, ';');
    const keywordInput = (this.domElement.querySelector('#keyword-input') as HTMLInputElement).value.replace(/,/g, ';');
    const subPracticeSelectOptions = getSelectedValues(this.domElement.querySelector('#sub-practice-select') as HTMLSelectElement).join(';');
    const productCategorySelectOptions = getSelectedValues(this.domElement.querySelector('#product-category-select') as HTMLSelectElement).join(';');
    const yearSelectOptions = getSelectedValues(this.domElement.querySelector('#year-select') as HTMLSelectElement).join(';');
    const clientSelectOptions = getSelectedValues(this.domElement.querySelector('#client-select') as HTMLSelectElement).join(';');
    */

    try {
      // Fetch the existing CSV file
      const response = await fetch(pathToAnalyticsCsvFile);
      if (!response.ok) {
        throw new Error('Failed to fetch the CSV file');
      }
      const csvText = await response.text();
  
      // Parse the CSV data
      const rows = csvText.split('\n');
      const newRow = [
        userEmail,
        searchInput,
        filter1Select,
        filter2Select,
        filter3Select,
        filter4Select,
        filter5Select,
        filter6Select,
        keywordInput,
        timeStamp,
        searchResultsNumber.toString()
      ];
      const updatedRows = [...rows, newRow.join(',')];
  
      // Convert the updated data back to CSV format
      const updatedCsvText = updatedRows.join('\n');
  
      // Write the updated CSV data back to the file
      try {
        const response = await fetch(pathToAnalyticsCsvFile, {
          method: 'PUT',
          headers: {
            'Content-Type': 'text/csv'
          },
          body: updatedCsvText
        });
    
        if (!response.ok) {
          throw new Error('Failed to write to the CSV file');
        }
        console.log('CSV file updated successfully');
      } catch (error) {
        console.error('Error writing to CSV file:', error);
      }
      console.log('New row added to the CSV file');
    } catch (error) {
      console.error('Error adding row to CSV file:', error);
    }
  }

  private async setSelectedOptions(filter1Select: string[], filter2Select: string[], filter3Select: string[], filter4Select: string[], filter5Select: string[], filter6Select: string[]): Promise<void> {
    const selectElementsMap: { [key: string]: string | string[] } = {
      'filter1-select': filter1Select,
      'filter2-select': filter2Select,
      'filter3-select': filter3Select,
      'filter4-select': filter4Select,
      'filter5-select': filter5Select,
      'filter6-select': filter6Select
    };
  
    const selectedElementsMap: { [key: string]: string } = {
      'filter1-select': 'filter1-selected',
      'filter2-select': 'filter2-selected',
      'filter3-select': 'filter2-selected',
      'filter4-select': 'filter4-selected',
      'filter5-select': 'filter5-selected',
      'filter6-select': 'filter6-selected'
    };

    // Mapping between selectElement IDs and corresponding div IDs
    const divMapping: { [key: string]: string } = {
      'filter1-select': 'filter1-dropdown-selected',
      'filter2-select': 'filter2-dropdown-selected',
      'filter3-select': 'filter3-dropdown-selected',
      'filter4-select': 'filter4-dropdown-selected',
      'filter5-select': 'filter5-dropdown-selected',
      'filter6-select': 'filter6-dropdown-selected'
    };

    //console.log("Setting Selected Options");
  
    Object.keys(selectElementsMap).forEach(selectId => {
      const selectElement = this.domElement.querySelector(`#${selectId}`) as HTMLSelectElement;
      const selectedValues = selectElementsMap[selectId];
      const selectedElementValue = selectedElementsMap[selectId];
      const selectedElement = this.domElement.querySelector(`#${selectedElementValue}`) as HTMLSelectElement;
  

      //console.log('Select Element:', selectElement);
      console.log('Selected Values:', selectedValues);
      //console.log('Selected Element:', selectedElement);
      //console.log('Selected element value:', selectedElementValue);

      if (selectElement && selectedElement) {
        selectedElement.innerHTML = '';
        
        Array.from(selectElement.options).forEach(option => {
          const newOption = document.createElement('option');
          newOption.value = option.value;
          if(option.value==='' && option.text==='Practice'){
            newOption.hidden = true; // Hide the option
          }
          newOption.text = option.text;
          selectedElement.add(newOption);
        });
  
        const correspondingDivId = divMapping[selectId];
        if (correspondingDivId) {
          const correspondingDiv = this.domElement.querySelector(`#${correspondingDivId}`) as HTMLDivElement;
          if (correspondingDiv) {
            const itemsDiv = correspondingDiv.nextElementSibling as HTMLDivElement;
            if (itemsDiv) {
              const items = itemsDiv.querySelectorAll('.selected-item');
              items.forEach(item => {
                const span = item.querySelector('span');
                if (span && !Array.from(selectedValues).some(opt => opt === span.textContent)) {
                  item.remove();
                }
              });
            }
          }
        }

        if (Array.isArray(selectedValues)) {
          selectedValues.forEach(selectedValue => {
            const optionToSelect = Array.from(selectedElement.options).find(option => option.value === selectedValue);
            console.log('Select Option ',selectedValue,' : ',optionToSelect);
            if (optionToSelect) {
              //optionToSelect.selected = true;
              
              // Get the corresponding div ID
              const correspondingDivId = divMapping[selectElement.id];
              console.log('Select element ID:', selectElement.id);
              console.log('Corresponding div ID:', correspondingDivId);
              if (correspondingDivId) {
                const correspondingDiv = document.getElementById(correspondingDivId);
                console.log('Corresponding div:', correspondingDiv);
                if (correspondingDiv) {
                  const matchingLabel = correspondingDiv.querySelector(`label[data-value="${selectedValue}"]`) as HTMLLabelElement;
                  console.log('Matching label:', matchingLabel);
                  if (matchingLabel) {
                    console.log('Clicking matching label');
                    matchingLabel.click();
                  }
                }
              }
            } else {
              console.error(`Option with value "${selectedValue}" not found in select element with ID "${selectId}".`);
            }
          });
        } else {
          const optionToSelect = Array.from(selectedElement.options).find(option => option.value === selectedValues);
          if (optionToSelect) {
            selectedElement.value = selectedValues as string;
          } else {
            console.error(`Option with value "${selectedValues}" not found in select element with ID "${selectId}".`);
          }
        }
      } else {
        console.error(`Select element with ID "${selectId}" not found.`);
      }
    });
  }
  
  private async searchSharePoint(): Promise<void> {
    this.searchResultsNumber = 0
    if (this.searchResultsNumber){
      const searchResultsHeading = document.getElementById('search-results-heading');
      if(searchResultsHeading){
        searchResultsHeading.innerHTML = `${this.homeIconHtml} Search Results for '${this.searchTerm}' (${this.searchResultsNumber})`;
      }
    }
    try {
      // Fetch the request digest - this is a security measure that tells SP that the request is valid and the user is authorized
      const digestResponse = await fetch(`${this.context.pageContext.web.absoluteUrl}/_api/contextinfo`, {
        method: 'POST',
        headers: {
          'accept': 'application/json;odata=verbose',
          'content-type': 'application/json;odata=verbose'
        }
      });

      if (!digestResponse.ok) {
        throw new Error(`HTTP error! status: ${digestResponse.status}`);
      }

      const digestData = await digestResponse.json();
      const requestDigest = digestData.d.GetContextWebInformation.FormDigestValue;

      // Retrieve selected filters
      const titleInput = (this.domElement.querySelector('#search-input') as HTMLInputElement).value;
      const filter1Select = (this.domElement.querySelector('#filter1-select') as HTMLSelectElement).selectedOptions;
      console.log("filter1Select", filter1Select);
      const filter2Select = (this.domElement.querySelector('#filter2-select') as HTMLSelectElement).selectedOptions;
      console.log("filter2Select", filter2Select);
      const filter3Select = (this.domElement.querySelector('#filter3-select') as HTMLSelectElement).selectedOptions;
      console.log("filter3Select", filter3Select);
      const filter4Select = (this.domElement.querySelector('#filter4-select') as HTMLSelectElement).selectedOptions;
      console.log("filter4Select", filter4Select);
      const filter5Select = (this.domElement.querySelector('#filter5-select') as HTMLSelectElement).selectedOptions;
      console.log("filter5Select", filter5Select);
      const filter6Select = (this.domElement.querySelector('#filter6-select') as HTMLSelectElement).selectedOptions;
      console.log("filter6Select", filter6Select);
      const keywordInput = (this.domElement.querySelector('#keyword-input') as HTMLInputElement).value;
      const searchInHeadlinesOnly = (this.domElement.querySelector('#filenamesonlycheckbox') as HTMLInputElement).checked;
      console.log("searchInHeadlinesOnly", searchInHeadlinesOnly);
      
      //console.log('filter2Select:', filter2Select);
      //console.log('filter3Select:', filter3Select);
      //console.log('filter4Select:', filter4Select);
      //console.log('filter5Select:', filter5Select);

      // Construct Querytext
      let queryText = ``;
      const searchInSites = ' AND(path:https://zeomega.sharepoint.com/sites/BizSolution OR path:https://zeomega.sharepoint.com/sites/MarketingCommunications OR path:https://zeomega.sharepoint.com/sites/TrainingResource OR path:https://zeomega.sharepoint.com/sites/ProductDocumentsRepository)'
      //if we want to search ONLY IN HEADLINES, then follow this syntax:
      if(searchInHeadlinesOnly){
        queryText += `(Title:${titleInput})`;
      }

      if (titleInput != "" && !searchInHeadlinesOnly) {
        console.log('Title input: ', titleInput);
        queryText += `${titleInput}* AND (ContentTypeId:0x0*) AND (dmsSearchContentStatus:Published) AND (IsContainer:false) AND (ContentClass:STS_ListItem_DocumentLibrary)${searchInSites}`;
      } else {
        queryText += ` AND (IsContainer:false) AND (dmsSearchContentStatus:Published) AND (ContentClass:STS_ListItem_DocumentLibrary)${searchInSites}`;
      }

      const hasMultipleFilter1Selected = filter1Select.length > 1;
      const hasMultipleFilter2Selected = filter2Select.length > 1;
      const hasMultipleFilter3Selected = filter3Select.length > 1;
      const hasMultipleFilter4Selected = filter4Select.length > 1;
      const hasMultipleFilter5Selected = filter5Select.length > 1;
      const hasMultipleFilter6Selected = filter6Select.length > 1;

      console.log('Multiple Departments Selected:', hasMultipleFilter1Selected);
      console.log('Multiple Product Names Selected:', hasMultipleFilter2Selected);
      console.log('Multiple Knowledge Category Selected:', hasMultipleFilter3Selected);
      console.log('Multiple Content Audience Selected:', hasMultipleFilter4Selected);
      console.log('Multiple Release Numbers Selected:', hasMultipleFilter5Selected);
      console.log('Multiple Modules Selected:', hasMultipleFilter6Selected);

      console.log('release number number of options:' , filter5Select.length);
      console.log('modules number of options:' , filter6Select.length);

      // Helper function to construct conditions
      const constructOrCondition = (selectedOptions: HTMLCollectionOf<HTMLOptionElement>, fieldName: string) => {
          const conditions = Array.from(selectedOptions)
              .filter(option => option.value !== '')
              .map(option => `(${fieldName}:${option.value})`);
          
          if (conditions.length === 1) {
              return ` AND ${conditions[0]}`;
          } else if (conditions.length > 1) {
              return ` AND (${conditions.join(' OR ')})`;
          } else {
              return '';
          }
      };

      if (filter1Select && filter1Select.length > 0) {
        queryText += constructOrCondition(filter1Select, 'dmsSearchDepartments');
      }
      if (filter2Select && filter2Select.length > 0) {
        queryText += constructOrCondition(filter2Select, 'dmsSearchProductName');
      }
      if (filter3Select && filter3Select.length > 0) {
        queryText += constructOrCondition(filter3Select, 'dmsSearchKnowledgeCategory');
      }
      if (filter4Select && filter4Select.length > 0) {
        queryText += constructOrCondition(filter4Select, 'dmsSearchContentAudience');
      }
      if (filter5Select && filter5Select.length > 0) {
        queryText += constructOrCondition(filter5Select, 'dmsSearchReleasenumber');
      }
      if (filter6Select && filter6Select.length > 0) {
        queryText += constructOrCondition(filter6Select, 'dmsSearchModule');
      }

      console.log('Query text:', queryText);
      console.log('Description Enabled in Results');
      const requestBody = {
        request: {
          __metadata: {
            type: 'Microsoft.Office.Server.Search.REST.SearchRequest'
          },
          Querytext: queryText, // Use the dynamically constructed query text
          SelectProperties: {
            results: [
              'Title', 'Author', 'Path', 'HitHighlightedSummary', 'DocumentSetDescription', 'FileExtension', 'IsDocument',
              'PictureThumbnailURL', 'FileType','LastModifiedTime','.spResourceUrl','.mediaBaseUrl','.thumbnailUrl','.callerStack','.correlationId','DefaultEncodingURL', 'dmsSearchDepartments','dmsSearchKnowledgeCategory','dmsSearchContentAudience','dmsSearchProjectName','dmsSearchContentStatus','dmsSearchContentOwnership','dmsSearchKeywords','dmsSearchProductName', 'dmsSearchReleasenumber','dmsSearchModule','dmsSearchClient','dmsSearchGeography','dmsSearchYear','IsFolder','IsContainer','ContentClass','Description'
            ]
          }
          ,RowLimit: 3000
        }
      };

      // Make the search request
      const response = await fetch(`${this.context.pageContext.web.absoluteUrl}/_api/search/postquery`, {
        method: 'POST',
        headers: {
          'accept': 'application/json;odata=verbose',
          'content-type': 'application/json;odata=verbose',
          'odata-version': '',
          'X-RequestDigest': requestDigest
        },
        body: JSON.stringify(requestBody)
      });

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const searchResults = await response.json();
      const filteredResults = keywordInput ? this.filterResultsByKeywords(searchResults.d.postquery.PrimaryQueryResult.RelevantResults.Table.Rows.results, keywordInput) : searchResults.d.postquery.PrimaryQueryResult.RelevantResults.Table.Rows.results;
      //console.log('Search results:', searchResults.d.postquery.PrimaryQueryResult.RelevantResults.Table.Rows.results);
      console.log('Filtered results:', filteredResults);
      console.log('Filtered results:', filteredResults.length);

      if (filteredResults.length === 0) {
        const searchResultsHeading = document.getElementById('search-results-heading');
        if(searchResultsHeading){
          searchResultsHeading.innerHTML = `Search Results (0)`;
          this.hideLoader();
        }
      } else {
        //Update the number of search results
        this.searchResultsNumber = filteredResults.length;
        if (this.searchResultsNumber){
          const searchResultsHeading = document.getElementById('search-results-heading');
          if(searchResultsHeading){
            searchResultsHeading.innerHTML = `${this.homeIconHtml} Search Results for '${this.searchTerm}' (${this.searchResultsNumber})`;
          }
        }
      }

      // Remove all event listeners from elements with ID 'downloadAsCsvBtn'
      //this.removeAllDownloadListenersById('downloadAsCsvBtn', 'click');
      //this.removeAllDownloadListenersById('downloadAsZipBtn', 'click');
      //console.log('Download buttons event listeners removed');

      //Render only the first 10 search results
      this.searchResultsToRender = filteredResults;
      await this.renderSearchResults(this.searchResultsToRender.slice(0, 10));
      this.searchesRendered = 10;
    } catch (error) {
      console.error('Error:', error);
    }
  }

  private filterResultsByKeywords(results: any[], keywords: string): any[] {
    if (!keywords) {
      return results;
    }
    
    // Split keywords by space or comma
    const keywordArray = keywords.split(/[\s,]+/).map(keyword => keyword.toLowerCase());

    return results.filter(row => {
      const cells = row.Cells.results;
      const dmsSearchKeywords = cells.find((cell: any) => cell.Key === 'dmsSearchKeywords')?.Value || '';
      const keywordsInMetadata = dmsSearchKeywords.toLowerCase().split(/[\s,]+/);

      // Check if any keyword in keywordArray is included in keywordsInMetadata
      return keywordArray.some(keyword => keywordsInMetadata.includes(keyword));
    }); 
  }

  // Helper function to convert a string to title case with exception words
  private toTitleCase(str: string): string {
    return str.split(' ').map((word) => {
      if (word.includes('_')) {
        return word.split('_').map((subWord) => {
            return subWord.charAt(0).toUpperCase() + subWord.substr(1).toLowerCase();
        }).join('_');
      } else {
          return word.charAt(0).toUpperCase() + word.substr(1).toLowerCase();
      }
    }).join(' ');
  }
    
  //Once the search results are retrieved from SharePoint, render them on the page for testing
  private async renderSearchResults(searchResults: any): Promise<void> {
    const searchResultsContainer = this.domElement.querySelector('#search-result') as HTMLDivElement;

    if (searchResultsContainer) {
      this.selectedDownloadItems = [];
      const processRow = async (row: any) => {
        const cells = row.Cells.results;
        let title = cells.find((cell: any) => cell.Key === 'Title').Value;
        let fileExtension = cells.find((cell: any) => cell.Key === 'FileExtension').Value;
        let fileType = cells.find((cell: any) => cell.Key === 'FileType').Value;
        let hitHighlightedSummary = cells.find((cell: any) => cell.Key === 'HitHighlightedSummary').Value;
        hitHighlightedSummary = hitHighlightedSummary.replace(/<c0>/g, '<mark>').replace(/<\/c0>/g, '</mark>');
        let pictureThumbnailURL = cells.find((cell: any) => cell.Key === 'PictureThumbnailURL')?.Value || 'https://zeomega.sharepoint.com/sites/zehub/SiteAssets/SPFx/Search/placeholder.jpg';
        const lastModifiedTime = cells.find((cell: any) => cell.Key === 'LastModifiedTime').Value;
        
        const dmsSearchDepartments = cells.find((cell: any) => cell.Key === 'dmsSearchDepartments')?.Value?.split('GP0')[0].slice(0, -1) || '';
        const dmsSearchKnowledgeCategory = cells.find((cell: any) => cell.Key === 'dmsSearchKnowledgeCategory')?.Value?.split('GP0')[0].slice(0, -1) || '';
        const dmsSearchContentAudience = cells.find((cell: any) => cell.Key === 'dmsSearchContentAudience')?.Value?.split('GP0')[0].slice(0, -1) || '';
        const dmsSearchProjectName = cells.find((cell: any) => cell.Key === 'dmsSearchProjectName')?.Value?.split('GP0')[0].slice(0, -1) || '';
        const dmsSearchContentStatus = cells.find((cell: any) => cell.Key === 'dmsSearchContentStatus')?.Value?.split('GP0')[0].slice(0, -1) || '';
        const dmsSearchContentOwnership = cells.find((cell: any) => cell.Key === 'dmsSearchContentOwnership')?.Value?.split('GP0')[0].slice(0, -1) || '';
        const dmsSearchKeywords = cells.find((cell: any) => cell.Key === 'dmsSearchKeywords')?.Value?.split('GP0')[0].slice(0, -1) || '';
        const dmsSearchProductName = cells.find((cell: any) => cell.Key === 'dmsSearchProductName')?.Value?.split('GP0')[0].slice(0, -1) || '';
        const dmsSearchReleasenumber = cells.find((cell: any) => cell.Key === 'dmsSearchReleasenumber')?.Value?.split('GP0')[0].slice(0, -1) || '';
        const dmsSearchModule = cells.find((cell: any) => cell.Key === 'dmsSearchModule')?.Value?.split('GP0')[0].slice(0, -1) || '';
        const dmsSearchClient = cells.find((cell: any) => cell.Key === 'dmsSearchClient')?.Value?.split('GP0')[0].slice(0, -1) || '';
        const dmsSearchGeography = cells.find((cell: any) => cell.Key === 'dmsSearchGeography')?.Value?.split('GP0')[0].slice(0, -1) || '';
        const dmsSearchYear = cells.find((cell: any) => cell.Key === 'dmsSearchYear')?.Value?.split('GP0')[0].slice(0, -1) || '';
        const description = cells.find((cell: any) => cell.Key === 'Description')?.Value || '';
        
        //let linkUrl = cells.find((cell: any) => cell.Key === 'OriginalPath').Value;
        let linkUrl = cells.find((cell: any) => cell.Key === 'DefaultEncodingURL').Value ? cells.find((cell: any) => cell.Key === 'DefaultEncodingURL').Value || '' : cells.find((cell: any) => cell.Key === 'Path').Value || '';
        let linkUrlWithoutQuery = cells.find((cell: any) => cell.Key === 'DefaultEncodingURL').Value ? cells.find((cell: any) => cell.Key === 'DefaultEncodingURL').Value || '' : cells.find((cell: any) => cell.Key === 'Path').Value || '';
      
        // Check if linkUrl contains ?web=1, if not append it
        if(linkUrl){
          if (!linkUrl.includes('?web=1')) {
            linkUrl += '?web=1';
          }
        }

        // Conver hitHighlightedSummary to title case with exception words
        hitHighlightedSummary = this.toTitleCase(hitHighlightedSummary);
        let itemDescription = ''
        if (description !== '') {
          itemDescription = description
        } else {
          itemDescription = hitHighlightedSummary
        }
      
        // Check if pictureThumbnailURL exists and is valid, if not set to default photo
        try {
          const response = await fetch(pictureThumbnailURL);
          if (!response.ok) {
            pictureThumbnailURL = 'https://zeomega.sharepoint.com/sites/zehub/SiteAssets/SPFx/Search/placeholder.jpg';
          }
        } catch (error) {
          pictureThumbnailURL = 'https://zeomega.sharepoint.com/sites/zehub/SiteAssets/SPFx/Search/placeholder.jpg';
        }
      
        // Parse the date string into a Date object
        const date = new Date(lastModifiedTime);
        // Format the date to dd-MM-yyyy hh:mm
        const formattedDate = date.toLocaleString('en-GB', {
          day: '2-digit',
          month: '2-digit',
          year: 'numeric',
          hour: '2-digit',
          minute: '2-digit',
          hour12: false
        });

        let titleWithExtension = title;
        if (!titleWithExtension.includes(fileExtension)) {
            if(fileExtension==='aspx'){
              titleWithExtension += `.${fileType}`;
            }else{
              titleWithExtension += `.${fileExtension}`;
            }   
        }

      // Function to check if the URL is a video file
      function isVideoFile(url: string): boolean {
        const videoExtensions = ['.mp4', '.webm', '.ogg'];
        const urlWithoutQuery = url.split('?')[0]; // Remove query parameters
        return videoExtensions.some(ext => urlWithoutQuery.toLowerCase().endsWith(ext));
      }

      // Function to get the MIME type of the video file
      function getVideoMimeType(url: string): string {
        const extension = url.split('.').pop()?.split('?')[0].toLowerCase();
        switch (extension) {
          case 'mp4':
            return 'video/mp4';
          case 'webm':
            return 'video/webm';
          case 'ogg':
            return 'video/ogg';
          default:
            return '';
        }
      }

      // Function to open video in a popup player
      function openVideoPlayer(url: string): void {
        const mimeType = getVideoMimeType(url);
        console.log(`Opening video: ${url}, MIME type: ${mimeType}`); // Debugging log
        const videoPlayerHtml = `
          <div id="videoPopup" style="position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); z-index: 1000; background: #000; padding: 10px;">
            <video id="videoPlayer" controls style="width: 100%; height: auto;">
              <source src="${url}" type="${mimeType}">
              Your browser does not support the video tag.
            </video>
            <button onclick="closeVideoPlayer()" style="position: absolute; top: 10px; right: 10px; background: #fff; border: none; padding: 5px;">Close</button>
          </div>
          <div id="videoOverlay" style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0, 0, 0, 0.5); z-index: 999;"></div>
        `;
        document.body.insertAdjacentHTML('beforeend', videoPlayerHtml);
      }

      // Function to close the video player
      function closeVideoPlayer(): void {
        const videoPopup = document.getElementById('videoPopup');
        const videoOverlay = document.getElementById('videoOverlay');
        if (videoPopup) videoPopup.remove();
        if (videoOverlay) videoOverlay.remove();
      }

      // Make the functions accessible globally
      (window as any).openVideoPlayer = openVideoPlayer;
      (window as any).closeVideoPlayer = closeVideoPlayer;

      let knowledgeCategory = dmsSearchKnowledgeCategory? dmsSearchKnowledgeCategory:'Nil';

      const itemHtml = `
        <div class="${styles.ItemResults}">
          <div class="${styles.ContentSection}">

            <!--
            <div class="${styles.CheckboxContainer}">
              <input type="checkbox" class="${styles.Checkbox}" data-title="${titleWithExtension}" data-linkurl="${linkUrl}" data-knowledgeCategory="${knowledgeCategory}">
            </div>
            -->
            <div class="${styles.ImgBox} test2">
              ${isVideoFile(linkUrl) ? `
                <a href="javascript:void(0);" onclick="openVideoPlayer('${linkUrlWithoutQuery}')" style="cursor: pointer;text-decoration:none;">
                  <img src="${pictureThumbnailURL}">
                </a>
              ` : `
                <a href="${linkUrl}" target="_blank" data-interception="off" style="cursor: pointer;text-decoration:none;">
                  <img src="${pictureThumbnailURL}">
                </a>
              `}
            </div>
            <div class="${styles.InnerContents}">
              <a href="${linkUrl}" target="_blank" data-interception="off" style="text-decoration:none;">
                  <h3>${title}</h3>
              </a>            
              <span>Uploaded - ${formattedDate}</span>
              
              <!--
              <div class="${styles.ratingContainer}">
                <div class="${styles.AvgRating}">
                  <div class="${styles.stars}">
                    <span class="${styles.star} ${styles.filled}">☆</span>
                    <span class="${styles.star} ${styles.filled}">☆</span>
                    <span class="${styles.star} ${styles.filled}">☆</span>
                    <span class="${styles.star} ${styles.filled}">☆</span>
                    <span class="${styles.star} ${styles.halfFilled}">☆</span>
                  </div>
                </div>
                <div class="${styles.averageRating}">
                  <span id="average-rating-value">4.3</span> out of 5
                </div>
              </div>
              -->

              <p class="${styles.Truncate}">${itemDescription}</p>
              <div class="${styles.Buttons}">
                ${dmsSearchDepartments ? `<a href="#">Departments: ${dmsSearchDepartments}</a>` : ''}
                ${dmsSearchKnowledgeCategory ? `<a href="#">Knowledge Category: ${dmsSearchKnowledgeCategory}</a>` : ''}
                ${dmsSearchContentAudience ? `<a href="#">Content Audience: ${dmsSearchContentAudience}</a>` : ''}
                ${dmsSearchProjectName ? `<a href="#">Project Name: ${dmsSearchProjectName}</a>` : ''}
                ${dmsSearchContentStatus ? `<a href="#">Content Status: ${dmsSearchContentStatus}</a>` : ''}
                ${dmsSearchContentOwnership ? `<a href="#">Content Ownership: ${dmsSearchContentOwnership}</a>` : ''}
                ${dmsSearchKeywords ? `<a href="#">Keywords: ${dmsSearchKeywords}</a>` : ''}
                ${dmsSearchProductName ? `<a href="#">Product Name: ${dmsSearchProductName}</a>` : ''}
                ${dmsSearchReleasenumber ? `<a href="#">Release Number: ${dmsSearchReleasenumber}</a>` : ''}
                ${dmsSearchModule ? `<a href="#">Module: ${dmsSearchModule}</a>` : ''}
                ${dmsSearchClient ? `<a href="#">Client: ${dmsSearchClient}</a>` : ''}
                ${dmsSearchGeography ? `<a href="#">Geography: ${dmsSearchGeography}</a>` : ''}
                ${dmsSearchYear ? `<a href="#">Year: ${dmsSearchYear}</a>` : ''}
              </div>
              
              <!--
              <div class="${styles.ratingSection}">
                <h5>Rate your experience</h5>
                <div class="${styles.rating}">
                  <input type="radio" id="star5" name="rating" value="5" />
                  <label for="star5" title="5 stars">☆</label>
                  <input type="radio" id="star4" name="rating" value="4" />
                  <label for="star4" title="4 stars">☆</label>
                  <input type="radio" id="star3" name="rating" value="3" />
                  <label for="star3" title="3 stars">☆</label>
                  <input type="radio" id="star2" name="rating" value="2" />
                  <label for="star2" title="2 stars">☆</label>
                  <input type="radio" id="star1" name="rating" value="1" />
                  <label for="star1" title="1 star">☆</label>
                </div>
              </div>
              -->
            </div>
          </div>
        </div>`;

          

          if (searchResultsContainer) {
            searchResultsContainer.innerHTML += itemHtml;
          }
      };

      for (const row of searchResults) {
        await processRow(row);
      }

    } else {
      console.error('Element with ID "search-results" not found.');
    }
  }

  //Function to jump to the search results after the search is performed
  private jumpToSearchResults(): void {
    const searchResultsHeading = document.getElementById('search-results-heading');
    if (searchResultsHeading) {
      searchResultsHeading.scrollIntoView({ behavior: 'smooth' });
    }
  }

  //Allows clicking on the home icon to clear all search results
  private addHomeIconEventHandler(): void {
    const homeIcon = document.getElementById('home-icon');
    const searchResultsContainer = this.domElement.querySelector('#search-result') as HTMLDivElement;
    const searchResults = this.domElement.querySelector('#search-results') as HTMLDivElement;

    if (homeIcon) {
      homeIcon.addEventListener('click', () => {
        searchResults.style.display = 'none';
        searchResultsContainer.innerHTML = ''; // Clear existing results
      });
    }
  }

  

  private addLoadMoreFunctionality(): void {
    if (this.searchesRendered< this.searchResultsNumber){
      const loadMoreBtn = this.domElement.querySelector('#load-more-btn') as HTMLButtonElement;
      const newLoadMoreBtn = loadMoreBtn.cloneNode(true) as HTMLButtonElement;

      if (loadMoreBtn.parentNode) {
        loadMoreBtn.parentNode.replaceChild(newLoadMoreBtn, loadMoreBtn);
      }
      loadMoreBtn.style.display = 'block';
      newLoadMoreBtn.style.display = 'block';

      newLoadMoreBtn.addEventListener('click', async () => {
        this.showLoader();
        this.searchesRendered += 10;
        await this.renderSearchResults(this.searchResultsToRender.slice(this.searchesRendered, this.searchesRendered + 10));
        this.hideLoader();
      });
    } else if (this.searchesRendered >= this.searchResultsNumber) {
      const loadMoreBtn = this.domElement.querySelector('#load-more-btn') as HTMLButtonElement;
      loadMoreBtn.style.display = 'none';
    }
  }
  //-------------------------------------------- END: All Funtions Related to Search -------------------------------------------//










  private populateSubPracticeContributeOptions(selectedPractice: string): void {
      const uniqueSubPractices = new Set<string>();
  
      console.log('Selected practice:', selectedPractice);
      if (selectedPractice && this.productNames[selectedPractice]) {
        this.productNames[selectedPractice].forEach(subPractice => {
          uniqueSubPractices.add(subPractice);
        });
        console.log('Unique sub-practices:', uniqueSubPractices);
      } else {
        // Populate with all sub-practices if no practice is selected
        Object.keys(this.productNames).forEach(practice => {
          this.productNames[practice].forEach(subPractice => {
            uniqueSubPractices.add(subPractice);
          });
        });
      }
      
    const subPracticeInputSelected = this.domElement.querySelector('#contribute-filter2-input') as HTMLSelectElement;
    subPracticeInputSelected.innerHTML = '<option value="" selected style="color: transparent; display: none;">Project</option>';
    
      uniqueSubPractices.forEach(subPractice => {
        const option = document.createElement('option');
        option.value = subPractice;
        option.text = subPractice;
        subPracticeInputSelected.appendChild(option);
      });
  }
  


  private populateProductCategoryContributeOptions(selectedPractice: string): void {
    const uniqueProductCategories = new Set<string>();
  
    if (selectedPractice && this.releaseNumbers[selectedPractice]) {
      this.releaseNumbers[selectedPractice].forEach(productCategory => {
        uniqueProductCategories.add(productCategory);
      });
    } else {
      // Populate with all product categories if no practice is selected
      Object.keys(this.releaseNumbers).forEach(practice => {
        //console.log("practice", practice);
        //console.log("practice !== Marketing", practice !== 'Marketing');
        if(practice !== 'Marketing'){
          this.releaseNumbers[practice].forEach(productCategory => {
            uniqueProductCategories.add(productCategory);
          });
        }
      });
    }
  
    const productCategoryInputSelect = this.domElement.querySelector('#contribute-filter3-input') as HTMLSelectElement;
    productCategoryInputSelect.innerHTML = '<option value="" selected style="color: transparent; display:none;">Document Type</option>';
    uniqueProductCategories.forEach(productCategory => {
      const option = document.createElement('option');
      option.value = productCategory;
      option.text = productCategory;
      productCategoryInputSelect.appendChild(option);
    });

  }

  private addTagsCloseIconListeners(): void {
    const closeIcons = this.domElement.querySelectorAll('.close-icon') as NodeListOf<HTMLElement>;
  
    closeIcons.forEach(icon => {
      icon.addEventListener('click', (event) => {
        const span = (event.target as HTMLElement).parentElement;
        if (span) {
          span.classList.add(styles.fadeOut);
          span.addEventListener('transitionend', () => {
            span.remove();
          }, { once: true });
        }
      });
    });
  }


  
  
    
  



  

  



  

  

  

  

  



  
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}