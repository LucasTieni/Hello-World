import { styled } from 'config/stitches.config'

export const Banner = styled('div', {
  width: '100%',
  height: '297px',
  fontSize: '56px',
  fontWeight: '500',
  display: 'flex',
  justifyContent: 'center',
  alignItems: 'center',
  backgroundSize: 'cover',
  backgroundPosition: 'center',
  backgroundRepeat: 'no-repeat',
  borderRadius: '8px 8px 0 0',
  backgroundColor: '#EFEFEF',
  '@bp3': {
    // fontSize: '32px',
    // height: '120px',
    display: 'none',
  },
})
