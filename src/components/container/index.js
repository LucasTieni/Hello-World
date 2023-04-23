import { styled } from 'config/stitches.config'

export const Container = styled('div', {
  width: '1200px',
  margin: '0 auto',
  paddingTop: '32px',
  '@bp1': {
    width: '100%',
    paddingTop: '12px',
  },
})

export const Grid = styled('div', {
  width: '100%',
  display: 'flex',
  flexWrap: 'wrap',
})

export const Col = styled('div', {
  width: 'calc(33% - 16px)',
  margin: '8px 8px',
  '.react-calendar ': {
    border: 'none',
  },
  '@bp3': {
    width: '100%',
  },
})

export const Markdown = styled('div', {
  width: '80%',
  margin: '0 auto',
  p: {
    marginTop: '32px',
  },
})
