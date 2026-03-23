import { WebPartContext } from '@microsoft/sp-webpart-base';
import { AnnouncementDataService } from '../services/AnnouncementDataService';

export interface IAnnouncementManagePanelProps {
  /** Whether the panel is open. */
  isOpen: boolean;
  /** SharePoint web part context for API calls. */
  context: WebPartContext;
  /** Reuse the data service instance from the parent carousel component. */
  dataService: AnnouncementDataService;
  /** Called when the panel is dismissed. dataChanged indicates if any CRUD operations occurred. */
  onDismiss: (dataChanged: boolean) => void;
}
