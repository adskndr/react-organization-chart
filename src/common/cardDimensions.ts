// Single source of truth for the person-card footprint. Anything that needs
// to size a box AROUND one or more cards (managerBox, leadershipBox, ...)
// should derive from these instead of hardcoding pixel numbers — otherwise
// the box size drifts out of sync the moment the card size changes, or the
// box only ever fits the number of cards it happened to be measured against
// (e.g. breaks as soon as a co-lead card is added next to the lead card).
export const PERSON_CARD_WIDTH = 340;
export const PERSON_CARD_HEIGHT = 73;
