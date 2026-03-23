import * as React from 'react';
import { IAnnouncementsCarouselProps } from './IAnnouncementsCarouselProps';
import { IAnnouncement, CelebrationType } from '../models/IAnnouncement';
import { AnnouncementDataService } from '../services/AnnouncementDataService';
import { CelebrationIconService } from '../services/CelebrationIconService';
import { ListProvisioningService } from '../services/ListProvisioningService';
import { PermissionService } from '../services/PermissionService';
import AnnouncementManagePanel from './AnnouncementManagePanel';
import { Spinner, SpinnerSize, MessageBar, MessageBarType, Icon, IconButton, DefaultButton } from '@fluentui/react';
import styles from './AnnouncementsCarousel.module.scss';

export interface IAnnouncementsCarouselState {
  announcements: IAnnouncement[];
  currentIndex: number;
  loading: boolean;
  error: string | null;
  isProvisioning: boolean;
  isPaused: boolean;
  isOwner: boolean;
  isManagePanelOpen: boolean;
}

export default class AnnouncementsCarousel extends React.Component<IAnnouncementsCarouselProps, IAnnouncementsCarouselState> {
  private dataService: AnnouncementDataService;
  private rotationTimer: number | null = null;

  constructor(props: IAnnouncementsCarouselProps) {
    super(props);

    this.state = {
      announcements: [],
      currentIndex: 0,
      loading: true,
      error: null,
      isProvisioning: false,
      isPaused: false,
      isOwner: false,
      isManagePanelOpen: false
    };

    this.dataService = new AnnouncementDataService(props.context, props.listName);
  }

  public async componentDidMount(): Promise<void> {
    const listReady = await this.ensureListExists();
    if (listReady) {
      await this.loadAnnouncements();
      this.startRotation();
    }

    const currentUserIsOwner = await PermissionService.checkIsOwner(this.props.context);
    this.setState({ isOwner: currentUserIsOwner });
  }

  public componentWillUnmount(): void {
    this.stopRotation();
  }

  public componentDidUpdate(prevProps: IAnnouncementsCarouselProps): void {
    if (prevProps.rotationInterval !== this.props.rotationInterval || prevProps.autoPlay !== this.props.autoPlay) {
      this.stopRotation();
      if (this.props.autoPlay && !this.state.isPaused) {
        this.startRotation();
      }
    }
  }

  private async ensureListExists(): Promise<boolean> {
    try {
      this.setState({ isProvisioning: true });
      await ListProvisioningService.ensureList(this.props.context);
      this.setState({ isProvisioning: false });
      return true;
    } catch (error) {
      console.error('Failed to provision list:', error);
      this.setState({
        error: `Failed to provision announcements list: ${error.message || 'Unknown error'}. Please check your permissions.`,
        isProvisioning: false,
        loading: false
      });
      return false;
    }
  }

  private async loadAnnouncements(): Promise<void> {
    try {
      this.setState({ loading: true, error: null });
      const announcements = await this.dataService.getActiveAnnouncements();

      if (announcements.length === 0) {
        this.setState({
          announcements: [],
          loading: false,
          error: null
        });
      } else {
        this.setState({
          announcements,
          loading: false,
          currentIndex: 0
        });
      }
    } catch (error) {
      console.error('Error loading announcements:', error);
      const errorMessage = error.message || 'Unknown error';
      this.setState({
        error: `Failed to load announcements: ${errorMessage}. The list may still be provisioning. Try refreshing the page in a few seconds.`,
        loading: false
      });
    }
  }

  private startRotation(): void {
    if (this.props.autoPlay && this.state.announcements.length > 1) {
      this.rotationTimer = setInterval(() => {
        this.goToNext();
      }, this.props.rotationInterval * 1000);
    }
  }

  private stopRotation(): void {
    if (this.rotationTimer) {
      clearInterval(this.rotationTimer);
      this.rotationTimer = null;
    }
  }

  private goToNext = (): void => {
    this.setState(prevState => ({
      currentIndex: (prevState.currentIndex + 1) % prevState.announcements.length
    }));
  };

  private goToPrevious = (): void => {
    this.setState(prevState => ({
      currentIndex: prevState.currentIndex === 0
        ? prevState.announcements.length - 1
        : prevState.currentIndex - 1
    }));
  };

  private goToSlide = (index: number): void => {
    this.setState({ currentIndex: index });
  };

  private handleOpenManagePanel = (): void => {
    this.setState({ isManagePanelOpen: true });
  };

  private handleManagePanelDismiss = (dataChanged: boolean): void => {
    this.setState({ isManagePanelOpen: false });
    if (dataChanged) {
      this.loadAnnouncements().catch(
        (reloadError) => console.error('Error reloading announcements after panel close:', reloadError)
      );
    }
  };

  private renderCelebrationIcon(announcement: IAnnouncement): JSX.Element | null {
    const showIcon = announcement.ShowCelebrationIcon != null ? announcement.ShowCelebrationIcon : true;
    if (!showIcon || announcement.CelebrationIcon === CelebrationType.None) {
      return null;
    }

    const iconSize = announcement.CelebrationIconSize || 60;

    if (announcement.CelebrationIcon === CelebrationType.Custom && announcement.CustomIconUrl) {
      const positionClass = `${styles.celebrationIcon} ${styles[announcement.CelebrationIconPosition || 'topRight']}`;

      return (
        <div
          className={positionClass}
          style={{
            width: `${iconSize + 20}px`,
            height: `${iconSize + 20}px`,
            overflow: 'hidden'
          }}
          title="Custom Icon"
        >
          <img
            src={announcement.CustomIconUrl}
            alt="Custom celebration icon"
            style={{
              width: '100%',
              height: '100%',
              objectFit: 'cover',
              borderRadius: '50%'
            }}
          />
        </div>
      );
    }

    const iconConfig = CelebrationIconService.getIconConfig(announcement.CelebrationIcon);
    if (!iconConfig) {
      return null;
    }

    const positionClass = `${styles.celebrationIcon} ${styles[announcement.CelebrationIconPosition || 'topRight']}`;
    const animationClass = iconConfig.animation ? styles[iconConfig.animation] : '';

    return (
      <div
        className={`${positionClass} ${animationClass}`}
        style={{
          backgroundColor: iconConfig.backgroundColor,
          color: iconConfig.color,
          fontSize: `${iconSize}px`,
          width: `${iconSize + 20}px`,
          height: `${iconSize + 20}px`
        }}
        title={iconConfig.label}
      >
        <span className={styles.celebrationEmoji}>{iconConfig.emoji}</span>
      </div>
    );
  }

  private renderAnnouncement(announcement: IAnnouncement): JSX.Element {
    const transitionClass = this.props.enableTransitions ? styles[this.props.transitionEffect] : '';

    // All visual properties come from the announcement itself (with defaults from data service mapping)
    const sizeMode = announcement.ImageSizeMode || 'fit';
    const imageHeight = announcement.ImageHeight || 400;
    const imageWidth = announcement.ImageWidth != null ? announcement.ImageWidth : 0;
    const imageFit = announcement.ImageFit || 'cover';
    const bgType = announcement.ImageBackgroundType || 'gradient';
    const solidBgColor = announcement.ImageBackgroundColor || '#667eea';
    const gradientStart = announcement.GradientStartColor || '#667eea';
    const gradientEnd = announcement.GradientEndColor || '#764ba2';
    const gradientDir = announcement.GradientDirection != null ? announcement.GradientDirection : 135;
    const overlayColor = announcement.OverlayGradientColor || '#000000';
    const overlayOpacity = announcement.OverlayOpacity != null ? announcement.OverlayOpacity : 30;
    const cardHeight = announcement.CardHeight || 500;
    const bgColor = announcement.BackgroundColor || '#ffffff';
    const borderRadius = announcement.BorderRadius != null ? announcement.BorderRadius : 8;
    const showShadow = announcement.ShowShadow != null ? announcement.ShowShadow : true;
    const titleColor = announcement.TitleColor || '#333333';
    const descColor = announcement.DescriptionColor || '#333333';

    // Image container background: solid or gradient
    const containerBackground = bgType === 'solid'
      ? solidBgColor
      : `linear-gradient(${gradientDir}deg, ${gradientStart} 0%, ${gradientEnd} 100%)`;
    const overlayGradient = `linear-gradient(to bottom, rgba(0, 0, 0, 0) 0%, ${overlayColor} 100%)`;

    // Image container style
    const imageContainerStyle: React.CSSProperties = {
      background: containerBackground
    };

    // Image style: CSS fit mode vs manual dimensions
    const imageStyle: React.CSSProperties = {};

    if (sizeMode === 'manual') {
      // Manual mode: explicit pixel dimensions
      imageContainerStyle.height = `${imageHeight}px`;
      if (imageWidth > 0) {
        imageStyle.width = `${imageWidth}px`;
        imageStyle.height = '100%';
      }
    } else {
      // Fit mode: CSS object-fit handles sizing
      imageContainerStyle.height = `${imageHeight}px`;
      imageStyle.objectFit = imageFit as any;
    }

    const hasRedirect = !!announcement.RedirectUrl;
    const redirectTarget = announcement.RedirectTarget || '_self';

    const slideContent = (
      <>
        {announcement.AnnouncementImage && (
          <div
            className={styles.imageContainer}
            style={imageContainerStyle}
          >
            <img
              src={announcement.AnnouncementImage}
              alt={announcement.Title}
              className={styles.announcementImage}
              style={imageStyle}
            />
            <div
              className={styles.imageOverlay}
              style={{ opacity: overlayOpacity / 100, background: overlayGradient }}
            />
            {this.renderCelebrationIcon(announcement)}
          </div>
        )}

        <div className={styles.contentContainer}>
          {this.props.showTitle && (
            <h2
              className={styles.announcementTitle}
              style={{
                fontSize: `${this.props.titleFontSize}px`,
                color: titleColor
              }}
            >
              {announcement.Title}
            </h2>
          )}

          {this.props.showDescription && (
            <div
              className={styles.announcementDescription}
              style={{
                fontSize: `${this.props.descriptionFontSize}px`,
                color: descColor
              }}
              dangerouslySetInnerHTML={{ __html: announcement.Description }}
            />
          )}
        </div>
      </>
    );

    const slideStyle: React.CSSProperties = {
      height: `${cardHeight}px`,
      backgroundColor: bgColor,
      borderRadius: `${borderRadius}px`,
      boxShadow: showShadow ? '0 4px 20px rgba(0,0,0,0.15)' : 'none'
    };

    if (hasRedirect) {
      return (
        <a
          href={announcement.RedirectUrl}
          target={redirectTarget}
          rel={redirectTarget === '_blank' ? 'noopener noreferrer' : undefined}
          className={`${styles.announcementSlide} ${styles.clickableSlide} ${transitionClass}`}
          style={{ ...slideStyle, textDecoration: 'none', color: 'inherit', display: 'flex', flexDirection: 'column' }}
        >
          {slideContent}
        </a>
      );
    }

    return (
      <div
        className={`${styles.announcementSlide} ${transitionClass}`}
        style={slideStyle}
      >
        {slideContent}
      </div>
    );
  }

  private renderNavigationDots(): JSX.Element | null {
    if (!this.props.showNavigationDots || this.state.announcements.length <= 1) {
      return null;
    }

    return (
      <div className={styles.navigationDots}>
        {this.state.announcements.map((_, index) => (
          <button
            key={index}
            className={`${styles.dot} ${index === this.state.currentIndex ? styles.activeDot : ''}`}
            onClick={() => this.goToSlide(index)}
            aria-label={`Go to slide ${index + 1}`}
          />
        ))}
      </div>
    );
  }

  private renderNavigationArrows(): JSX.Element | null {
    if (!this.props.showNavigationArrows || this.state.announcements.length <= 1) {
      return null;
    }

    return (
      <>
        <IconButton
          className={`${styles.navButton} ${styles.prevButton}`}
          iconProps={{ iconName: 'ChevronLeft' }}
          onClick={this.goToPrevious}
          aria-label="Previous announcement"
        />
        <IconButton
          className={`${styles.navButton} ${styles.nextButton}`}
          iconProps={{ iconName: 'ChevronRight' }}
          onClick={this.goToNext}
          aria-label="Next announcement"
        />
      </>
    );
  }

  private renderManageButton(): JSX.Element | null {
    if (!this.state.isOwner) return null;

    return (
      <div className={styles.manageButtonContainer}>
        <DefaultButton
          text="Manage Announcements"
          iconProps={{ iconName: 'Settings' }}
          onClick={this.handleOpenManagePanel}
          className={styles.manageButton}
          ariaLabel="Open announcements management panel"
        />
        <AnnouncementManagePanel
          isOpen={this.state.isManagePanelOpen}
          context={this.props.context}
          dataService={this.dataService}
          onDismiss={this.handleManagePanelDismiss}
        />
      </div>
    );
  }

  public render(): React.ReactElement<IAnnouncementsCarouselProps> {
    const { loading, error, announcements, currentIndex, isProvisioning } = this.state;

    if (isProvisioning) {
      return (
        <div className={styles.announcementsCarousel}>
          <div className={styles.centerContent}>
            <Spinner size={SpinnerSize.large} label="Setting up announcements list..." />
          </div>
        </div>
      );
    }

    if (loading) {
      return (
        <div className={styles.announcementsCarousel}>
          <div className={styles.centerContent}>
            <Spinner size={SpinnerSize.large} label="Loading announcements..." />
          </div>
        </div>
      );
    }

    if (error) {
      return (
        <div className={styles.announcementsCarousel}>
          {this.renderManageButton()}
          <MessageBar messageBarType={MessageBarType.error}>
            {error}
          </MessageBar>
        </div>
      );
    }

    if (announcements.length === 0) {
      const bgType = this.props.emptyStateBackgroundType || 'gradient';
      const emptyStateStyle: React.CSSProperties = {};

      if (bgType === 'solid') {
        emptyStateStyle.background = this.props.emptyStateBackgroundColor || '#f5f7fa';
      } else if (bgType === 'image') {
        const imgUrl = this.props.emptyStateBackgroundImage;
        if (imgUrl) {
          emptyStateStyle.backgroundImage = `url('${imgUrl}')`;
          emptyStateStyle.backgroundSize = 'cover';
          emptyStateStyle.backgroundPosition = 'center';
        }
      } else {
        const start = this.props.emptyStateGradientStart || '#f5f7fa';
        const end = this.props.emptyStateGradientEnd || '#e4e8ee';
        const dir = this.props.emptyStateGradientDirection != null ? this.props.emptyStateGradientDirection : 135;
        emptyStateStyle.background = `linear-gradient(${dir}deg, ${start} 0%, ${end} 100%)`;
      }

      return (
        <div className={styles.announcementsCarousel}>
          {this.renderManageButton()}
          <div className={styles.emptyState} style={emptyStateStyle}>
            <div
              className={styles.emptyStateIconWrapper}
              style={{ color: this.props.emptyStateIconColor }}
            >
              <Icon
                iconName={this.props.emptyStateIcon}
                className={styles.emptyStateIcon}
              />
            </div>
            <p
              className={styles.emptyStateMessage}
              style={{ color: this.props.emptyStateTextColor || '#444444' }}
            >
              {this.props.emptyStateMessage}
            </p>
          </div>
        </div>
      );
    }

    const currentAnnouncement = announcements[currentIndex];

    return (
      <div className={styles.announcementsCarousel}>
        {this.renderManageButton()}
        <div className={styles.carouselContainer}>
          {this.renderAnnouncement(currentAnnouncement)}
          {this.renderNavigationArrows()}
          {this.renderNavigationDots()}
        </div>
      </div>
    );
  }
}
