/**
 * Standalone Angular component wrapping PptxViewer. Copy this file into your
 * application (Angular 16+). Usage:
 *
 *   <app-pptx-viewer [src]="url" [options]="{ mode: 'slideshow' }"
 *                    (loaded)="onLoad($event)" (failed)="onError($event)"></app-pptx-viewer>
 */
import {
  AfterViewInit,
  ChangeDetectionStrategy,
  Component,
  ElementRef,
  EventEmitter,
  Input,
  NgZone,
  OnChanges,
  OnDestroy,
  Output,
  SimpleChanges,
  ViewChild,
} from "@angular/core";
import { PptxViewer, type PptxError, type PptxSource, type RenderedPresentation, type ViewerOptions } from "@meshesha/pptxjs";

@Component({
  selector: "app-pptx-viewer",
  standalone: true,
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `<div #host class="pptx-host"></div>`,
  styles: [":host { display: block; }"],
})
export class PptxViewerComponent implements AfterViewInit, OnChanges, OnDestroy {
  @Input() src: PptxSource | null = null;
  /** Read when the component is created; use `viewer.setOptions()` afterwards. */
  @Input() options: ViewerOptions = {};
  @Output() loaded = new EventEmitter<RenderedPresentation>();
  @Output() failed = new EventEmitter<PptxError>();
  @Output() slideChange = new EventEmitter<number>();

  @ViewChild("host", { static: true }) private host!: ElementRef<HTMLDivElement>;
  /** The underlying viewer (available after view init). */
  viewer: PptxViewer | null = null;

  constructor(private readonly zone: NgZone) {}

  ngAfterViewInit(): void {
    // run outside Angular so animation frames and timers do not trigger change detection
    this.zone.runOutsideAngular(() => {
      const v = new PptxViewer(this.host.nativeElement, this.options);
      v.on("load", ({ presentation }) => this.zone.run(() => this.loaded.emit(presentation)));
      v.on("error", ({ error }) => this.zone.run(() => this.failed.emit(error)));
      v.on("slidechange", ({ index }) => this.zone.run(() => this.slideChange.emit(index)));
      this.viewer = v;
      if (this.src) v.load(this.src).catch(() => undefined);
    });
  }

  ngOnChanges(changes: SimpleChanges): void {
    if (changes["src"] && !changes["src"].firstChange && this.viewer && this.src) {
      const src = this.src;
      this.zone.runOutsideAngular(() => this.viewer!.load(src).catch(() => undefined));
    }
  }

  ngOnDestroy(): void {
    this.viewer?.destroy();
    this.viewer = null;
  }
}
