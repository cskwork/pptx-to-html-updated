(function() {
    const slideWidth = 960.00;
    const slideHeight = 720.00;
    const slides = Array.from(document.querySelectorAll('.slide'));
    const stage = document.querySelector('.slide-stage');
    const progress = document.getElementById('progress');
    const totalSlides = slides.length;
    const prevBtn = document.getElementById('prev');
    const nextBtn = document.getElementById('next');
    const blackout = document.getElementById('transitionBlackout');
    const slideTransitions = [{"effect": "cut", "duration": 0, "delay": 0, "direction": null, "orientation": null, "throughBlack": false, "spokes": null, "shape": null, "speed": null, "advanceOnClick": true, "advanceAfter": null}];
    const defaultTransition = {
        effect: 'cut',
        duration: 0,
        delay: 0,
        direction: null,
        orientation: null,
        throughBlack: false,
        spokes: null,
        shape: null,
        speed: null,
        advanceOnClick: true,
        advanceAfter: null
    };
    const cssVarKeys = [
        '--ppt-transition-duration',
        '--ppt-transition-delay',
        '--ppt-translate-in-x',
        '--ppt-translate-in-y',
        '--ppt-translate-out-x',
        '--ppt-translate-out-y',
        '--ppt-scale-in',
        '--ppt-scale-out',
        '--ppt-rotate-in',
        '--ppt-rotate-out',
        '--ppt-mask-start',
        '--ppt-mask-angle',
        '--ppt-mask-shape',
        '--ppt-clip-inset-top',
        '--ppt-clip-inset-right',
        '--ppt-clip-inset-bottom',
        '--ppt-clip-inset-left',
        '--ppt-wheel-spokes',
        '--ppt-clip-shape'
    ];
    let currentSlide = 0;
    let isTransitioning = false;
    let autoAdvanceTimer = null;

    function applyScale() {
        if (!stage || !stage.parentElement) {
            return;
        }
        const wrapper = stage.parentElement;
        const scaleX = wrapper.clientWidth / slideWidth;
        const scaleY = wrapper.clientHeight / slideHeight;
        const scale = Math.min(scaleX, scaleY);
        stage.style.transform = `scale(${scale})`;
    }

    function mergeTransition(index) {
        const raw = slideTransitions[index] || {};
        const config = Object.assign({}, defaultTransition);
        Object.keys(raw).forEach(key => {
            if (raw[key] !== undefined && raw[key] !== null) {
                config[key] = raw[key];
            }
        });
        if (!config.spokes && config.effect === 'wheel') {
            config.spokes = 6;
        }
        return config;
    }

    function invertDirection(direction) {
        switch (direction) {
            case 'l': return 'r';
            case 'r': return 'l';
            case 'u': return 'd';
            case 'd': return 'u';
            case 'tl': return 'br';
            case 'tr': return 'bl';
            case 'bl': return 'tr';
            case 'br': return 'tl';
            default: return direction;
        }
    }

    function directionVector(direction) {
        switch (direction) {
            case 'l':
                return { inX: '-100%', inY: '0%', outX: '100%', outY: '0%', maskAngle: '90deg', maskStart: '-125%', clip: { top: '0%', right: '0%', bottom: '0%', left: '100%' }, rotate: '90deg' };
            case 'r':
                return { inX: '100%', inY: '0%', outX: '-100%', outY: '0%', maskAngle: '270deg', maskStart: '125%', clip: { top: '0%', right: '100%', bottom: '0%', left: '0%' }, rotate: '-90deg' };
            case 'u':
                return { inX: '0%', inY: '-100%', outX: '0%', outY: '100%', maskAngle: '0deg', maskStart: '-125%', clip: { top: '100%', right: '0%', bottom: '0%', left: '0%' }, rotate: '90deg' };
            case 'd':
                return { inX: '0%', inY: '100%', outX: '0%', outY: '-100%', maskAngle: '180deg', maskStart: '125%', clip: { top: '0%', right: '0%', bottom: '100%', left: '0%' }, rotate: '-90deg' };
            case 'tl':
                return { inX: '-100%', inY: '-100%', outX: '100%', outY: '100%', maskAngle: '45deg', maskStart: '-125%', clip: { top: '100%', right: '0%', bottom: '0%', left: '100%' }, rotate: '90deg' };
            case 'tr':
                return { inX: '100%', inY: '-100%', outX: '-100%', outY: '100%', maskAngle: '315deg', maskStart: '125%', clip: { top: '100%', right: '100%', bottom: '0%', left: '0%' }, rotate: '-90deg' };
            case 'bl':
                return { inX: '-100%', inY: '100%', outX: '100%', outY: '-100%', maskAngle: '135deg', maskStart: '-125%', clip: { top: '0%', right: '0%', bottom: '100%', left: '100%' }, rotate: '90deg' };
            case 'br':
                return { inX: '100%', inY: '100%', outX: '-100%', outY: '-100%', maskAngle: '225deg', maskStart: '125%', clip: { top: '0%', right: '100%', bottom: '100%', left: '0%' }, rotate: '-90deg' };
            default:
                return { inX: '0%', inY: '0%', outX: '0%', outY: '0%', maskAngle: '0deg', maskStart: '125%', clip: { top: '0%', right: '0%', bottom: '0%', left: '0%' }, rotate: '90deg' };
        }
    }

    function setClip(target, top, right, bottom, left) {
        if (!target) {
            return;
        }
        target.style.setProperty('--ppt-clip-inset-top', top);
        target.style.setProperty('--ppt-clip-inset-right', right);
        target.style.setProperty('--ppt-clip-inset-bottom', bottom);
        target.style.setProperty('--ppt-clip-inset-left', left);
    }

    function shapeClipFor(shape) {
        switch ((shape || '').toLowerCase()) {
            case 'circle':
                return 'circle(0% at 50% 50%)';
            case 'diamond':
                return 'polygon(50% 0%, 100% 50%, 50% 100%, 0% 50%)';
            case 'plus':
                return 'polygon(40% 0%, 60% 0%, 60% 40%, 100% 40%, 100% 60%, 60% 60%, 60% 100%, 40% 100%, 40% 60%, 0% 60%, 0% 40%, 40% 40%)';
            case 'star':
                return 'polygon(50% 0%, 61% 35%, 98% 35%, 68% 57%, 79% 91%, 50% 70%, 21% 91%, 32% 57%, 2% 35%, 39% 35%)';
            case 'square':
                return 'inset(0% 0% 0% 0%)';
            default:
                return 'circle(0% at 50% 50%)';
        }
    }

    function clearTransitionClasses(slide) {
        if (!slide) {
            return;
        }
        slide.classList.remove('ppt-transition-enter', 'ppt-transition-exit', 'pre-active');
        Array.from(slide.classList).forEach(cls => {
            if (cls.startsWith('ppt-transition-')) {
                slide.classList.remove(cls);
            }
        });
    }

    function clearTransitionVariables(slide) {
        if (!slide) {
            return;
        }
        cssVarKeys.forEach(key => slide.style.removeProperty(key));
    }

    function getTransitionClasses(effect) {
        switch (effect) {
            case 'cover':
                return { inClass: 'ppt-transition-cover-in', outClass: null };
            case 'uncover':
                return { inClass: null, outClass: 'ppt-transition-uncover-out' };
            case 'flash':
                return { inClass: 'ppt-transition-flash-in', outClass: 'ppt-transition-flash-out' };
            default:
                return { inClass: `ppt-transition-${effect}-in`, outClass: `ppt-transition-${effect}-out` };
        }
    }

    function applyEffectVariables(incoming, outgoing, config, vectors, forward) {
        if (incoming) {
            incoming.style.setProperty('--ppt-transition-duration', `${config.duration}ms`);
            incoming.style.setProperty('--ppt-transition-delay', `${config.delay}ms`);
        }
        if (outgoing) {
            outgoing.style.setProperty('--ppt-transition-duration', `${config.duration}ms`);
        }

        switch (config.effect) {
            case 'push':
                if (incoming) {
                    incoming.style.setProperty('--ppt-translate-in-x', vectors.inX);
                    incoming.style.setProperty('--ppt-translate-in-y', vectors.inY);
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-translate-out-x', vectors.outX);
                    outgoing.style.setProperty('--ppt-translate-out-y', vectors.outY);
                }
                break;
            case 'cover':
                if (incoming) {
                    incoming.style.setProperty('--ppt-translate-in-x', vectors.inX);
                    incoming.style.setProperty('--ppt-translate-in-y', vectors.inY);
                }
                break;
            case 'uncover':
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-translate-out-x', vectors.outX);
                    outgoing.style.setProperty('--ppt-translate-out-y', vectors.outY);
                }
                if (incoming) {
                    incoming.classList.remove('pre-active');
                    incoming.classList.add('active');
                }
                break;
            case 'wipe':
                if (incoming) {
                    setClip(incoming, vectors.clip.top, vectors.clip.right, vectors.clip.bottom, vectors.clip.left);
                }
                if (outgoing) {
                    setClip(outgoing, '0%', '0%', '0%', '0%');
                }
                break;
            case 'split': {
                const orient = (config.orientation || '').toLowerCase();
                if (incoming) {
                    if (orient === 'horz') {
                        setClip(incoming, '50%', '0%', '50%', '0%');
                    } else {
                        setClip(incoming, '0%', '50%', '0%', '50%');
                    }
                }
                if ((config.direction || 'in').toLowerCase() === 'out' && outgoing) {
                    if (orient === 'horz') {
                        setClip(outgoing, '50%', '0%', '50%', '0%');
                    } else {
                        setClip(outgoing, '0%', '50%', '0%', '50%');
                    }
                }
                break;
            }
            case 'bars':
            case 'blinds':
            case 'comb':
                if (incoming) {
                    incoming.style.setProperty('--ppt-mask-angle', vectors.maskAngle);
                    incoming.style.setProperty('--ppt-mask-start', vectors.maskStart);
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-mask-angle', vectors.maskAngle);
                    outgoing.style.setProperty('--ppt-mask-start', vectors.maskStart);
                }
                break;
            case 'strips':
                if (incoming) {
                    incoming.style.setProperty('--ppt-mask-angle', vectors.maskAngle);
                    incoming.style.setProperty('--ppt-mask-start', vectors.maskStart);
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-mask-angle', vectors.maskAngle);
                    outgoing.style.setProperty('--ppt-mask-start', vectors.maskStart);
                }
                break;
            case 'checker':
                if (incoming) {
                    incoming.style.setProperty('--ppt-mask-start', '220%');
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-mask-start', '220%');
                }
                break;
            case 'zoom':
                if (incoming) {
                    incoming.style.setProperty('--ppt-scale-in', forward ? '0.35' : '1.15');
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-scale-out', forward ? '1.1' : '0.35');
                }
                break;
            case 'wheel':
                if (incoming) {
                    incoming.style.setProperty('--ppt-wheel-spokes', `${config.spokes || 6}`);
                    incoming.style.setProperty('--ppt-mask-angle', vectors.maskAngle);
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-wheel-spokes', `${config.spokes || 6}`);
                    outgoing.style.setProperty('--ppt-mask-angle', vectors.maskAngle);
                }
                break;
            case 'shape':
                if (incoming) {
                    incoming.style.setProperty('--ppt-clip-shape', shapeClipFor(config.shape));
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-clip-shape', shapeClipFor(config.shape));
                }
                break;
            case 'gallery':
            case 'flip':
            case 'cube':
                if (incoming) {
                    incoming.style.setProperty('--ppt-rotate-in', forward ? '-75deg' : '75deg');
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-rotate-in', forward ? '75deg' : '-75deg');
                }
                break;
            case 'switch':
                if (incoming) {
                    incoming.style.setProperty('--ppt-rotate-in', forward ? '-55deg' : '55deg');
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-rotate-in', forward ? '55deg' : '-55deg');
                }
                break;
            case 'conveyor':
                if (incoming) {
                    incoming.style.setProperty('--ppt-translate-in-x', vectors.inX);
                    incoming.style.setProperty('--ppt-translate-in-y', vectors.inY);
                    incoming.style.setProperty('--ppt-rotate-in', forward ? '-25deg' : '25deg');
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-translate-out-x', vectors.outX);
                    outgoing.style.setProperty('--ppt-translate-out-y', vectors.outY);
                    outgoing.style.setProperty('--ppt-rotate-in', forward ? '25deg' : '-25deg');
                }
                break;
            case 'pan':
                if (incoming) {
                    incoming.style.setProperty('--ppt-translate-in-x', vectors.inX);
                    incoming.style.setProperty('--ppt-translate-in-y', vectors.inY);
                    incoming.style.setProperty('--ppt-scale-in', '1.1');
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-scale-out', '0.85');
                }
                break;
            case 'rotate':
                if (incoming) {
                    incoming.style.setProperty('--ppt-rotate-in', forward ? '-120deg' : '120deg');
                }
                if (outgoing) {
                    outgoing.style.setProperty('--ppt-rotate-in', forward ? '120deg' : '-120deg');
                }
                break;
            case 'box':
                if (incoming) {
                    setClip(incoming, '50%', '25%', '50%', '25%');
                }
                break;
            default:
                break;
        }
    }

    function activateBlackout(active, duration) {
        if (!blackout) {
            return;
        }
        blackout.style.setProperty('--ppt-transition-duration', `${duration}ms`);
        blackout.classList.toggle('active', active);
    }

    function finalizeTransition(previousIndex, nextIndex, config) {
        const outgoing = slides[previousIndex];
        const incoming = slides[nextIndex];

        clearTransitionClasses(outgoing);
        clearTransitionClasses(incoming);
        clearTransitionVariables(outgoing);
        clearTransitionVariables(incoming);

        if (outgoing) {
            outgoing.classList.remove('active');
            outgoing.classList.remove('pre-active');
        }
        if (incoming) {
            incoming.classList.remove('pre-active');
            incoming.classList.add('active');
        }

        if (config.throughBlack) {
            window.setTimeout(() => activateBlackout(false, Math.max(200, Math.round(config.duration / 2))), 20);
        }

        currentSlide = nextIndex;
        isTransitioning = false;
        updateControls();
        initializeCharts();
        scheduleAutoAdvance(currentSlide);
        if (progress) {
            progress.style.width = `${((currentSlide + 1) / totalSlides) * 100}%`;
        }
        document.dispatchEvent(new Event('slideChanged'));
    }

    function playTransition(nextIndex, config, meta) {
        isTransitioning = true;
        clearTimeout(autoAdvanceTimer);

        const previousIndex = currentSlide;
        const outgoing = slides[previousIndex];
        const incoming = slides[nextIndex];
        const vectors = directionVector(config.direction);

        if (incoming) {
            clearTransitionClasses(incoming);
            clearTransitionVariables(incoming);
            incoming.classList.add('ppt-transition-enter');
            if (config.effect !== 'uncover') {
                incoming.classList.add('pre-active');
            } else {
                incoming.classList.add('active');
            }
        }

        if (outgoing) {
            clearTransitionClasses(outgoing);
            clearTransitionVariables(outgoing);
            outgoing.classList.add('ppt-transition-exit');
            outgoing.classList.add('active');
        }

        applyEffectVariables(incoming, outgoing, config, vectors, meta.forward);

        const classes = getTransitionClasses(config.effect);
        window.requestAnimationFrame(() => {
            if (incoming && classes.inClass) {
                incoming.classList.add(classes.inClass);
            }
            if (outgoing && classes.outClass) {
                outgoing.classList.add(classes.outClass);
            }
        });

        if (config.throughBlack) {
            activateBlackout(true, Math.max(200, Math.round(config.duration / 2)));
        }

        const totalDuration = config.duration + config.delay;
        window.setTimeout(() => finalizeTransition(previousIndex, nextIndex, config), totalDuration + 48);
    }

    function performImmediate(index) {
        clearTimeout(autoAdvanceTimer);
        const previousIndex = currentSlide;
        const outgoing = slides[previousIndex];
        const incoming = slides[index];

        if (outgoing) {
            clearTransitionClasses(outgoing);
            clearTransitionVariables(outgoing);
            outgoing.classList.remove('active');
            outgoing.classList.remove('pre-active');
        }

        if (incoming) {
            clearTransitionClasses(incoming);
            clearTransitionVariables(incoming);
            incoming.classList.add('active');
            incoming.classList.remove('pre-active');
        }

        currentSlide = index;
        isTransitioning = false;
        updateControls();
        initializeCharts();
        scheduleAutoAdvance(currentSlide);
        if (progress) {
            progress.style.width = `${((currentSlide + 1) / totalSlides) * 100}%`;
        }
        document.dispatchEvent(new Event('slideChanged'));
    }

    function scheduleAutoAdvance(index) {
        clearTimeout(autoAdvanceTimer);
        const cfg = mergeTransition(index);
        if (cfg.advanceAfter && cfg.advanceAfter > 0) {
            autoAdvanceTimer = window.setTimeout(() => {
                if (index === currentSlide) {
                    const nextIndex = Math.min(totalSlides - 1, index + 1);
                    if (nextIndex !== index) {
                        showSlide(nextIndex, { force: true });
                    }
                }
            }, cfg.advanceAfter);
        }
    }

    function updateControls() {
        if (prevBtn) {
            prevBtn.disabled = currentSlide === 0;
        }
        if (nextBtn) {
            const cfg = mergeTransition(currentSlide);
            nextBtn.disabled = currentSlide === totalSlides - 1 || cfg.advanceOnClick === false;
        }
        if (progress) {
            progress.style.width = `${((currentSlide + 1) / totalSlides) * 100}%`;
        }
    }

    function initializeCharts() {
        if (typeof Chart === 'undefined') {
            return;
        }
        document.querySelectorAll('canvas[data-chart-config]').forEach(canvas => {
            if (canvas.dataset.initialized === '1') {
                return;
            }
            try {
                const config = JSON.parse(canvas.dataset.chartConfig);
                new Chart(canvas.getContext('2d'), config);
                canvas.dataset.initialized = '1';
            } catch (error) {
                console.error('Chart initialization failed', error);
            }
        });
    }

    function showSlide(index, options = {}) {
        if (index < 0 || index >= totalSlides || index === currentSlide) {
            return;
        }
        if (isTransitioning) {
            return;
        }
        const forward = index > currentSlide;
        const currentConfig = mergeTransition(currentSlide);
        if (forward && !options.force && currentConfig.advanceOnClick === false) {
            return;
        }

        const nextConfig = mergeTransition(index);
        if (!forward && nextConfig.direction) {
            nextConfig.direction = invertDirection(nextConfig.direction);
        }

        if (nextConfig.effect === 'cut') {
            performImmediate(index);
            return;
        }

        if (!nextConfig.duration || nextConfig.duration < 1) {
            nextConfig.duration = 700;
        }

        playTransition(index, nextConfig, { forward });
    }

    if (prevBtn) {
        prevBtn.addEventListener('click', () => showSlide(currentSlide - 1, { force: true }));
    }
    if (nextBtn) {
        nextBtn.addEventListener('click', () => showSlide(currentSlide + 1));
    }

    document.addEventListener('keydown', event => {
        if (event.key === 'ArrowRight' || event.key === 'PageDown' || event.key === ' ') {
            if (mergeTransition(currentSlide).advanceOnClick === false) {
                return;
            }
            event.preventDefault();
            showSlide(Math.min(totalSlides - 1, currentSlide + 1));
        } else if (event.key === 'ArrowLeft' || event.key === 'PageUp') {
            event.preventDefault();
            showSlide(Math.max(0, currentSlide - 1), { force: true });
        }
    });

    window.addEventListener('resize', () => window.requestAnimationFrame(applyScale));

    window.addEventListener('load', () => {
        slides.forEach((slide, idx) => {
            slide.classList.toggle('active', idx === 0);
            slide.classList.toggle('pre-active', idx !== 0);
        });
        updateControls();
        applyScale();
        initializeCharts();
        scheduleAutoAdvance(0);
        document.dispatchEvent(new Event('slideChanged'));
    });

    window.showSlide = showSlide;
    window.nextSlide = () => showSlide(currentSlide + 1);
    window.prevSlide = () => showSlide(currentSlide - 1, { force: true });
})();