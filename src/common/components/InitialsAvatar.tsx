import * as React from 'react';
import styles from './InitialsAvatar.module.scss';
import { getInitials, pickGradient } from '../util/format';

export interface IInitialsAvatarProps {
  /** Pre-computed initials; if omitted they are derived from `name`. */
  initials?: string;
  name?: string;
  /** Diameter in px (default 42). */
  size?: number;
  /** Explicit CSS gradient; if omitted one is derived from name/initials. */
  gradient?: string;
}

export const InitialsAvatar: React.FunctionComponent<IInitialsAvatarProps> = (props) => {
  const size: number = props.size || 42;
  const text: string = props.initials || getInitials(props.name);
  const gradient: string = props.gradient || pickGradient(props.name || text);

  return (
    <div
      className={styles.avatar}
      aria-hidden="true"
      style={{
        width: size,
        height: size,
        background: gradient,
        fontSize: Math.round(size * 0.34)
      }}
    >
      {text}
    </div>
  );
};
