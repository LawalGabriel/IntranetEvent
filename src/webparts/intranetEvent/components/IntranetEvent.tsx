/* eslint-disable no-void */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import { useState, useEffect, useRef, useCallback } from 'react';
import styles from './IntranetEvent.module.scss';
import type { IIntranetEventProps, IIntranetEvent } from './IIntranetEventProps';
import { SPFx, spfi } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Placeholder } from '@pnp/spfx-controls-react';
import { Icon } from '@fluentui/react/lib/Icon';

const IntranetEvent: React.FC<IIntranetEventProps> = (props) => {
  const [eventItems, setEventItems] = useState<IIntranetEvent[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const spRef = useRef<any>(null);
  const maxRows = props.maxRows || 4;
  const rowHeight = props.rowHeight || '70px';
  const webPartTitle = props.webPartTitle || 'EVENTS';

  const loadEventItems = useCallback(async () => {
    try {
      setIsLoading(true);
      setErrorMessage('');

      const items: IIntranetEvent[] = await spRef.current.web.lists
        .getByTitle(props.listTitle)
        .items
        .select(
          "Id",
          "Title",
          "EventDate",
          "EndDate",
          "Location",
          "Description",
          "Category"
        )
        .orderBy("EventDate", true)();

      setEventItems(items);
      console.log('Loaded event items:', items);
      setIsLoading(false);
    } catch (error: any) {
      console.error('Error loading event items:', error);
      setIsLoading(false);
      setErrorMessage(`Failed to load event items. Please check if the list "${props.listTitle}" exists and you have permissions. Error: ${error.message}`);
    }
  }, [props.listTitle]);

  useEffect(() => {
    spRef.current = spfi().using(SPFx(props.context));
    void loadEventItems();
  }, [props.listTitle, props.context, loadEventItems]);

  // Format date to get day, month, and year
  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  const formatEventDate = (dateString: string) => {
    if (!dateString) return { day: '', month: '', year: '' };
    
    const date = new Date(dateString);
    const day = ('0' + date.getDate()).slice(-2);
    const month = date.toLocaleString('en-US', { month: 'short' }).toUpperCase();
    const year = date.getFullYear().toString();
    
    return { day, month, year };
  };

  // Format time range
  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  const formatTimeRange = (startDate: string, endDate: string) => {
    if (!startDate || !endDate) return '';
    
    const start = new Date(startDate);
    const end = new Date(endDate);
    
    // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
    const formatTime = (date: Date) => {
      let hours = date.getHours();
      const minutes = date.getMinutes();
      const ampm = hours >= 12 ? 'pm' : 'am';
      hours = hours % 12;
      hours = hours ? hours : 12;
      const minutesStr = minutes < 10 ? '0' + minutes : minutes;
      return `${hours}:${minutesStr}${ampm}`;
    };

    return `${formatTime(start)}-${formatTime(end)}`;
  };

  // Calculate container height based on maxRows
  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  const calculateContainerHeight = () => {
    const heightValue = parseInt(rowHeight);
    const unit = rowHeight.replace(heightValue.toString(), '');
    return (maxRows * heightValue) + unit;
  };

  if (isLoading) {
    return (
      <div className={styles.intranetEvent}>
        <div className={styles.loading}>Loading events...</div>
      </div>
    );
  }

  if (errorMessage) {
    return (
      <div className={styles.intranetEvent}>
        <Placeholder
          iconName='Error'
          iconText='Error'
          description={errorMessage}
        >
          <button onClick={() => loadEventItems()}>
            Retry
          </button>
        </Placeholder>
      </div>
    );
  }


  return (
    <div className={styles.intranetEvent}>
      {/* Header - Apply ALL styles directly to h1 */}
      <div 
        className={styles.header} 
        style={{ 
          backgroundColor: props.headerBgColor || '#2c3e50',
        }}
      >
        <h1 style={{ 
          color: props.headerTextColor || '#ffffff',
          fontSize: props.headerFontSize || 'clamp(18px, 2vw, 24px)',
          fontWeight: props.headerFontWeight || '600'
        }}>
          <Icon 
            iconName="Calendar" 
            className={styles.headerIcon}
            style={{ 
              color: props.headerTextColor || '#ffffff',
              fontSize: '0.8em' // Make icon slightly smaller than text
            }}
          />
          {webPartTitle}
        </h1>
      </div>

      {/* Events Container */}
      <div 
        className={styles.eventsContainer}
        style={{ maxHeight: calculateContainerHeight() }}
      >
        {eventItems.length === 0 ? (
          <div className={styles.noEvents}>
            <Icon iconName="Calendar" className={styles.noEventsIcon} />
            No upcoming events
          </div>
        ) : (
          eventItems.slice(0, maxRows).map((item: IIntranetEvent, index: number) => {
            const { day, month, year } = formatEventDate(item.EventDate);
            const timeRange = formatTimeRange(item.EventDate, item.EndDate);
            
            // FIX: Ensure eventBgColorAlt is properly used
            const rowBgColor = index % 2 === 0 
              ? (props.eventBgColor || '#ffffff')
              : (props.eventBgColorAlt || '#f8f9fa');
            
            const rowTextColor = props.eventTextColor || '#333333';
            const timeLocationColor = props.timeLocationColor || '#666666';
            
            return (
              <div 
                key={item.Id} 
                className={styles.eventItem}
                style={{
                  backgroundColor: rowBgColor, // This should now work
                  minHeight: rowHeight
                }}
              >
                {/* Date Box */}
                <div 
                  className={styles.dateBox}
                  style={{
                    backgroundColor: props.dateBgColor || '#e74c3c',
                    color: props.dateTextColor || '#ffffff'
                  }}
                >
                  <div 
                    className={styles.day}
                    style={{ color: props.dateTextColor || '#ffffff' }}
                  >
                    {day}
                  </div>
                  <div className={styles.monthYear}>
                    <span 
                      className={styles.month}
                      style={{ color: props.dateTextColor || '#ffffff' }}
                    >
                      {month}
                    </span>
                    <span 
                      className={styles.year}
                      style={{ color: props.dateTextColor || '#ffffff' }}
                    >
                      {year}
                    </span>
                  </div>
                </div>

                {/* Event Details */}
                <div className={styles.eventDetails}>
                  <div className={styles.eventTopRow}>
                    <div 
                      className={styles.eventTitle}
                      style={{ color: rowTextColor }}
                    >
                      {item.Title}
                    </div>
                    {item.Category && (
                      <div 
                        className={styles.eventCategory}
                        style={{
                          backgroundColor: props.categoryBgColor || '#3498db',
                          color: props.categoryTextColor || '#ffffff'
                        }}
                      >
                        {item.Category}
                      </div>
                    )}
                  </div>
                  
                  <div className={styles.eventBottomRow}>
                    {(timeRange || item.Location) && (
                      <>
                        {timeRange && (
                          <div 
                            className={styles.eventTimeLocation}
                            style={{ color: timeLocationColor }}
                          >
                            <Icon 
                              iconName="Clock" 
                              className={styles.timeIcon}
                              style={{ color: timeLocationColor }}
                            />
                            <span>{timeRange}</span>
                          </div>
                        )}
                        
                        {item.Location && (
                          <div 
                            className={styles.eventTimeLocation}
                            style={{ color: timeLocationColor }}
                          >
                            <Icon 
                              iconName="MapPin" 
                              className={styles.locationIcon}
                              style={{ color: timeLocationColor }}
                            />
                            <span>{item.Location}</span>
                          </div>
                        )}
                      </>
                    )}
                  </div>
                </div>
              </div>
            );
          })
        )}
      </div>
    </div>
  );
};


export default IntranetEvent;