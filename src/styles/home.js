import { styled } from 'config/stitches.config'

export const Card = styled('div', {
  width: '100%',
  background: '$neutral1',
  padding: '22px',
  borderRadius: '8px',
  border: '1px #C9C9C9 solid',
  borderLeft: '4px #CACACA solid',
})

export const Title = styled('div', {
  fontWeight: '600',
  fontSize: '18px',
  textTransform: 'capitalize',
  '@bp3': {
    fontSize: '16px',
  },
})

export const Links = styled('div', {
  fontWeight: '400',
  fontSize: '16px',
  color: '#000',
  marginTop: '12px',
  display: 'flex',
  alignItems: 'center',
  cursor: 'pointer',
  img: {
    marginRight: '8px',
  },
  svg: {
    marginRight: '8px',
  },
  '@bp3': {
    fontSize: '14px',
  },
})

export const highlight = styled('div', {
  color: 'green',
})
