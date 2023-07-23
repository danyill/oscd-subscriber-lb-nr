import { msg } from '@lit/localize';
import { css, html, LitElement, TemplateResult } from 'lit';
import { property, query } from 'lit/decorators.js';

import '@material/mwc-button';
import '@material/mwc-dialog';
import '@material/mwc-formfield';
import '@material/mwc-switch';

import type { Dialog } from '@material/mwc-dialog';
import type { Switch } from '@material/mwc-switch';
import { EditEvent, isUpdate, newEditEvent } from '@openscd/open-scd-core';

import { subscribe, unsubscribe } from '@openenergytools/scl-lib';
import {
  findControlBlock,
  findFCDAs,
  isSubscribed,
} from './foundation/subscription/subscription.js';
// import { findFCDAs } from './foundation/subscription/subscription.js';

/**
 * Checks that two FCDAs are identical except the second has a quality
 * attribute.
 *
 * @param a - an SCL FCDA element.
 * @param b - an SCL FCDA element.
 * @returns A boolean indicating that they are a pair.
 */
function isfcdaPairWithQuality(a: Element, b: Element): boolean {
  return ['ldInst', 'prefix', 'lnClass', 'lnInst', 'doName', 'daName'].every(
    attr =>
      a.getAttribute(attr) === b.getAttribute(attr) ||
      (attr === 'daName' &&
        b.getAttribute('daName')?.split('.').slice(-1)[0] === 'q')
  );
}

/**
 * Match a value/quality pair for the internal address on a NR device.
 *
 * A typical example might be:
 *   * RxExtIn1;/Ind/stVal
 *   * RxExtIn1;/Ind/q
 *
 * @param a - an ExtRef intAddr for a NR device.
 * @param b - an ExtRef intAddr for a NR device.
 * @returns a boolean indicating if the intAddr suggests these
 * match except in the data attribute.
 */
function extRefMatchNR(a: Element, b: Element): boolean {
  const aParts = a.getAttribute('intAddr')?.split('.');
  const bParts = b.getAttribute('intAddr')?.split('.');

  // if missing an intAddr then not a match
  if (!aParts || !bParts) return false;

  return (
    JSON.stringify(aParts?.slice(0, aParts.length - 1)) ===
      JSON.stringify(bParts?.slice(0, bParts.length - 1)) &&
    bParts[bParts.length - 1].slice(-1) === 'q'
  );
}

function shouldListen(event: Event): boolean {
  const initiatingTarget = <Element>event.composedPath()[0];
  return (
    initiatingTarget instanceof Element &&
    initiatingTarget.getAttribute('identity') ===
      'danyill.oscd-subscriber-later-binding' &&
    initiatingTarget.hasAttribute('allowexternalplugins')
  );
}

export default class SubscriberLaterBindingNR extends LitElement {
  /** The document being edited as provided to plugins by [[`OpenSCD`]]. */
  @property({ attribute: false })
  doc!: XMLDocument;

  @property({ attribute: false })
  docName!: string;

  preEventExtRef: (Element | null)[] = [];

  ignoreSupervision: boolean = false;

  @query('#dialog') dialogUI?: Dialog;

  @query('#enabled') enabledUI?: Switch;

  @property({ attribute: false })
  enabled: boolean = localStorage.getItem('oscd-subscriber-lb-nr') === 'true';

  constructor() {
    super();

    // record information to capture intention
    window.addEventListener(
      'oscd-edit',
      event => this.captureMetadata(event as EditEvent),
      { capture: true }
    );

    window.addEventListener('oscd-edit', event => {
      if (shouldListen(event)) this.modifyAdditionalExtRefs(event as EditEvent);
    });
  }

  async run(): Promise<void> {
    if (this.dialogUI) this.dialogUI.show();
  }

  /**
   * This method records the ExtRefs prior to the EditEvent and
   * also records whether supervisions can be changed for later
   * processing.
   * @param event - An EditEvent.
   */
  protected captureMetadata(event: EditEvent): void {
    if (shouldListen(event)) {
      const initiatingTarget = <Element>event.composedPath()[0];
      // is the later binding subscriber plugin allowing supervisions
      this.ignoreSupervision =
        initiatingTarget.hasAttribute('ignoresupervision') ?? false;

      // Infinity as 1 due to error type instantiation error
      // https://github.com/microsoft/TypeScript/issues/49280
      const flatEdits = [event.detail].flat(Infinity as 1);

      this.preEventExtRef = flatEdits.map(edit => {
        if (isUpdate(edit) && edit.element.tagName === 'ExtRef')
          return this.doc.importNode(edit.element, true);
        return null;
      });
    }
  }

  /**
   * Assess ExtRef for being associate with GOOSE value/quality and
   * dispatch subscribe or unsubscribe events.
   *
   * @param firstExtRef - an ExtRef subject to subscribe/unsubscribe.
   * @param preEventExtRef - an ExtRef subject to subscribe/unsubscribe.
   * but prior to the evnet.
   * @param firstFcda - the matching FCDA to the first ExtRef.
   * @returns
   */
  protected modifyValueAndQualityPair(
    firstExtRef: Element,
    preEventExtRef: Element | null,
    firstFcda: Element
  ): void {
    const controlBlock = findControlBlock(firstExtRef);

    // Else match value/quality pairs
    const nextFcda = firstFcda.nextElementSibling;
    const nextExtRef = firstExtRef.nextElementSibling;

    // They must exist
    if (!nextFcda || !nextExtRef) return;

    const wasSubscribed = preEventExtRef && isSubscribed(preEventExtRef);
    if (
      extRefMatchNR(firstExtRef, nextExtRef) &&
      nextFcda &&
      isfcdaPairWithQuality(firstFcda, nextFcda)
    ) {
      if (!wasSubscribed && isSubscribed(firstExtRef) && controlBlock)
        this.dispatchEvent(
          newEditEvent(
            subscribe({
              sink: nextExtRef,
              source: { fcda: nextFcda, controlBlock },
            })
          )
        );

      if (wasSubscribed && !isSubscribed(firstExtRef))
        this.dispatchEvent(newEditEvent(unsubscribe([nextExtRef])));
    }
  }

  /**
   * Will generate and dispatch further EditEvents based on matching an
   * ExtRef with subsequent ExtRefs and the first FCDA with subsequent
   * FCDAs. Uses both `extRef` and `preEventExtRef` to ensure subscription
   * information is available for unsubscribe edits.
   * @param extRef - an SCL ExtRef element
   * @param preEventExtRef - an SCL ExtRef element cloned before changes
   * @returns
   */
  protected processNRExtRef(extRef: Element, preEventExtRef: Element | null) {
    // look for change in subscription pre and post-event
    if (
      !isSubscribed(extRef) &&
      preEventExtRef &&
      !isSubscribed(preEventExtRef)
    )
      return;

    const fcdas = isSubscribed(extRef)
      ? findFCDAs(extRef)
      : findFCDAs(preEventExtRef!);

    let firstFcda: Element | undefined;
    // eslint-disable-next-line prefer-destructuring
    if (fcdas) firstFcda = fcdas[0];

    // must be able to locate the first fcda to continue
    if (!firstFcda) return;

    // If we have a value/quality pair do that
    this.modifyValueAndQualityPair(extRef, preEventExtRef, firstFcda);
  }

  /**
   * Either subscribe or unsubscribe from additional ExtRefs adjacent
   * to any ExtRefs found within an event if conditions are met for
   * manufacturer and event type.
   *
   * Assumes that all adding and removing of subscriptions is done
   * through Update edits of ExtRef elements.
   *
   * Only looks at IEDs whose manufacturer is "NRR"
   *
   * @param event - An open-scd-core EditEvent
   * @returns nothing.
   */
  protected modifyAdditionalExtRefs(event: EditEvent): void {
    if (!this.enabled) return;

    // Infinity as 1 due to error type instantiation error
    // https://github.com/microsoft/TypeScript/issues/49280
    const flatEdits = [event.detail].flat(Infinity as 1);

    flatEdits.forEach((edit, index) => {
      if (
        isUpdate(edit) &&
        edit.element.tagName === 'ExtRef' &&
        edit.element?.closest('IED')?.getAttribute('manufacturer') === 'NRR'
      ) {
        this.processNRExtRef(edit.element, this.preEventExtRef[index]);
      }
    });

    // restore pre-event cached data
    this.preEventExtRef = [];
    this.ignoreSupervision = false;
  }

  // TODO: Update URL when subscriber later binding is shepherded by OpenSCD organisation
  render(): TemplateResult {
    return html`<mwc-dialog
      id="dialog"
      heading="${msg('Subscriber Later Binding - NR')}"
    >
      <p>${msg('This plugin works with the')}
        <a
          href="https://github.com/danyill/oscd-subscriber-later-binding"
          target="_blank"
          >Subscriber Later Binding plugin</a
        >
        ${msg('to provide enhancements for NR Electric devices:')}
        <ul>
          <li>${msg('Automatic quality mapping')}</li>
        </ul>
        ${msg('for subscribing and unsubscribing.')}
      </p>
      <mwc-formfield label="${msg('Enabled')}">
      <!-- TODO: Remove ?checked when open-scd uses later version of mwc-components -->
        <mwc-switch id="enabled" ?selected=${this.enabled} ?checked=${
      this.enabled
    }>
        </mwc-switch>
      </mwc-formfield>
      <mwc-button
        label="${msg('Close')}"
        slot="primaryAction"
        icon="done"
        @click="${() => {
          // TODO: Remove when open-scd uses later version of mwc-components.
          this.enabled =
            this.enabledUI!.selected ?? (<any>this.enabledUI!).checked ?? false;
          localStorage.setItem('oscd-subscriber-lb-nr', `${this.enabled}`);
          if (this.dialogUI) this.dialogUI.close();
        }}"
      ></mwc-button>
    </mwc-dialog>`;
  }

  static styles = css`
    mwc-formfield {
      float: right;
    }
  `;
}
