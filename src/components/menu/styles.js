import { styled } from 'config/stitches.config'

export const Menu = styled('div', {
  width: '100%',
  height: '64px',
  background: '$primary1',
  padding: '16px',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  ul: {
    padding: '0',
    li: {
      display: 'inline-block',
      a: {
        padding: '0 16px',
        color: '$neutral1',
      },
    },
  },
  '@bp3': {
    display: 'none',
  },
})
