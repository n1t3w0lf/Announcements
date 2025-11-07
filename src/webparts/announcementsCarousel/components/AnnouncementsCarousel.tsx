import * as React from 'react';
import { IAnnouncementsCarouselProps } from './IAnnouncementsCarouselProps';
import { IAnnouncement, CelebrationType } from '../models/IAnnouncement';
import { AnnouncementDataService } from '../services/AnnouncementDataService';
import { CelebrationIconService } from '../services/CelebrationIconService';
import { ListProvisioningService } from '../services/ListProvisioningService';
import { Spinner, SpinnerSize, MessageBar, MessageBarType, IconButton } from '@fluentui/react';
import styles from './AnnouncementsCarousel.module.scss';

export interface IAnnouncementsCarouselState {
  announcements: IAnnouncement[];
  currentIndex: number;
  loading: boolean;
  error: string | null;
  isProvisioning: boolean;
  isPaused: boolean;
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
      isPaused: false
    };

    this.dataService = new AnnouncementDataService(props.context, props.listName);
  }

  public async componentDidMount(): Promise<void> {
    const listReady = await this.ensureListExists();
    if (listReady) {
      await this.loadAnnouncements();
      this.startRotation();
    }
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

  private togglePause = (): void => {
    const newPausedState = !this.state.isPaused;
    this.setState({ isPaused: newPausedState });

    if (newPausedState) {
      this.stopRotation();
    } else {
      this.startRotation();
    }
  };

  private renderCelebrationIcon(announcement: IAnnouncement): JSX.Element | null {
    if (!this.props.showCelebrationIcon || announcement.CelebrationIcon === CelebrationType.None) {
      return null;
    }

    const iconConfig = CelebrationIconService.getIconConfig(announcement.CelebrationIcon);
    if (!iconConfig) {
      return null;
    }

    const positionClass = `${styles.celebrationIcon} ${styles[announcement.CelebrationIconPosition || 'top-right']}`;
    const animationClass = iconConfig.animation ? styles[iconConfig.animation] : '';

    return (
      <div
        className={`${positionClass} ${animationClass}`}
        style={{
          backgroundColor: iconConfig.backgroundColor,
          color: iconConfig.color,
          fontSize: `${this.props.celebrationIconSize}px`,
          width: `${this.props.celebrationIconSize + 20}px`,
          height: `${this.props.celebrationIconSize + 20}px`
        }}
        title={iconConfig.label}
      >
        <span className={styles.celebrationEmoji}>{iconConfig.emoji}</span>
      </div>
    );
  }

  private renderAnnouncement(announcement: IAnnouncement): JSX.Element {
    const transitionClass = this.props.enableTransitions ? styles[this.props.transitionEffect] : '';

    return (
      <div
        className={`${styles.announcementSlide} ${transitionClass}`}
        style={{
          height: `${this.props.height}px`,
          backgroundColor: this.props.backgroundColor,
          borderRadius: `${this.props.borderRadius}px`,
          boxShadow: this.props.showShadow ? '0 4px 20px rgba(0,0,0,0.15)' : 'none'
        }}
      >
        {announcement.AnnouncementImage && (
          <div className={styles.imageContainer}>
            <img
              src={announcement.AnnouncementImage}
              alt={announcement.Title}
              className={styles.announcementImage}
            />
            <div
              className={styles.imageOverlay}
              style={{ opacity: this.props.overlayOpacity / 100 }}
            />
            {this.renderCelebrationIcon(announcement)}
          </div>
        )}

        <div className={styles.contentContainer} style={{ color: this.props.textColor }}>
          {this.props.showTitle && (
            <h2
              className={styles.announcementTitle}
              style={{ fontSize: `${this.props.titleFontSize}px` }}
            >
              {announcement.Title}
            </h2>
          )}

          {this.props.showDescription && (
            <div
              className={styles.announcementDescription}
              style={{ fontSize: `${this.props.descriptionFontSize}px` }}
              dangerouslySetInnerHTML={{ __html: announcement.Description }}
            />
          )}

          <div className={styles.dateInfo}>
            <span className={styles.dateLabel}>
              Valid: {announcement.ValidFrom.toLocaleDateString()} - {announcement.ValidTo.toLocaleDateString()}
            </span>
          </div>
        </div>
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

  private renderPlayPauseButton(): JSX.Element | null {
    if (!this.props.autoPlay || this.state.announcements.length <= 1) {
      return null;
    }

    return (
      <IconButton
        className={styles.playPauseButton}
        iconProps={{ iconName: this.state.isPaused ? 'Play' : 'Pause' }}
        onClick={this.togglePause}
        aria-label={this.state.isPaused ? 'Play carousel' : 'Pause carousel'}
      />
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
          <MessageBar messageBarType={MessageBarType.error}>
            {error}
          </MessageBar>
        </div>
      );
    }

    if (announcements.length === 0) {
      return (
        <div className={styles.announcementsCarousel}>
          <MessageBar messageBarType={MessageBarType.info}>
            No active announcements to display. Add announcements to the "{this.props.listName}" list to get started.
          </MessageBar>
        </div>
      );
    }

    const currentAnnouncement = announcements[currentIndex];

    return (
      <div className={styles.announcementsCarousel}>
        <div className={styles.carouselContainer}>
          {this.renderAnnouncement(currentAnnouncement)}
          {this.renderNavigationArrows()}
          {this.renderNavigationDots()}
          {this.renderPlayPauseButton()}
        </div>
      </div>
    );
  }
}
