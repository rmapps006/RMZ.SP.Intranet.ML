import * as React from 'react';
import styles from './ContentManager.module.scss';
import { Icon } from '@fluentui/react';

interface IRichTextFieldProps {
  label: string;
  value: string;
  // Changes when the edited item switches, so the editor re-seeds its HTML.
  resetKey: string | number;
  onChange: (html: string) => void;
}

interface IToolButton {
  cmd: string;
  icon: string;
  title: string;
  value?: string;
}

const BUTTONS: IToolButton[] = [
  { cmd: 'bold', icon: 'Bold', title: 'Bold' },
  { cmd: 'italic', icon: 'Italic', title: 'Italic' },
  { cmd: 'underline', icon: 'Underline', title: 'Underline' },
  { cmd: 'insertUnorderedList', icon: 'BulletedList', title: 'Bulleted list' },
  { cmd: 'insertOrderedList', icon: 'NumberedList', title: 'Numbered list' }
];

/**
 * A minimal, dependency-free rich-text editor built on a contentEditable div and
 * document.execCommand. Emits HTML for storage (rendered on the news/benefit
 * detail pages). Seeds its content when `resetKey` changes so switching items or
 * opening the add form starts from the right value without fighting the caret.
 */
const RichTextField: React.FunctionComponent<IRichTextFieldProps> = (props) => {
  const ref = React.useRef<HTMLDivElement>(null);

  // Only re-seed when the edited item changes (resetKey), not on every keystroke,
  // so the caret isn't reset while typing.
  React.useEffect(() => {
    if (ref.current) {
      ref.current.innerHTML = props.value || '';
    }
  }, [props.resetKey]);

  const emit = (): void => {
    if (ref.current) {
      props.onChange(ref.current.innerHTML);
    }
  };

  const exec = (cmd: string, value?: string): void => {
    try {
      if (ref.current) {
        ref.current.focus();
      }
      // execCommand is deprecated but still the simplest cross-browser way to do
      // basic inline formatting for a lightweight editor.
      document.execCommand(cmd, false, value);
      emit();
    } catch {
      /* ignore unsupported command */
    }
  };

  const addLink = (): void => {
    // eslint-disable-next-line no-alert
    const url: string | null = typeof window !== 'undefined' && window.prompt ? window.prompt('Link URL (https://…)') : null;
    if (url) {
      exec('createLink', url);
    }
  };

  return (
    <div className={styles.rtField}>
      <label className={styles.fieldLabel}>{props.label}</label>
      <div className={styles.rtToolbar}>
        {BUTTONS.map((b) => (
          <button
            key={b.cmd}
            type="button"
            className={styles.rtBtn}
            title={b.title}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => exec(b.cmd, b.value)}
          >
            <Icon iconName={b.icon} />
          </button>
        ))}
        <button
          type="button"
          className={styles.rtBtn}
          title="Insert link"
          onMouseDown={(e) => e.preventDefault()}
          onClick={addLink}
        >
          <Icon iconName="Link" />
        </button>
        <button
          type="button"
          className={styles.rtBtn}
          title="Clear formatting"
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => exec('removeFormat')}
        >
          <Icon iconName="ClearFormatting" />
        </button>
      </div>
      <div
        ref={ref}
        className={styles.rtEditor}
        contentEditable
        role="textbox"
        aria-multiline="true"
        aria-label={props.label}
        onInput={emit}
        onBlur={emit}
      />
    </div>
  );
};

export default RichTextField;
