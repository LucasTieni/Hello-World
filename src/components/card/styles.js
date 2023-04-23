import { styled } from 'config/stitches.config'

export const Card = styled('div', {
  width: '100%',
  background: '$neutral1',
  padding: '22px',
  borderRadius: '8px',
  border: '1px #C9C9C9 solid',
  borderLeft: '4px #CACACA solid',
})

export const Thumb = styled('div', {
  width: '100%',
  height: '200px',
  backgroundSize: 'cover',
  backgroundPosition: 'center',
  backgroundRepeat: 'no-repeat',
  borderRadius: '8px 8px 0 0',
  // border: '1px solid #E9E9E9',
  // background: '#D9D9D9',
  marginBottom: '16px',
  // display: 'flex',
  // justifyContent: 'center',
  // alignItems: 'center',
})

export const Empty = styled('div', {
  width: '24px',
  height: '24px',
})

export const Title = styled('div', {
  fontWeight: '600',
  fontSize: '18px',
  textTransform: 'capitalize',
  '@bp3': {
    fontSize: '16px',
  },
})

export const Date = styled('div', {
  fontWeight: '400',
  fontSize: '16px',
  color: '#8F8F8F',
  marginTop: '4px',
  '@bp3': {
    fontSize: '14px',
  },
})

export const Description = styled('div', {
  fontWeight: '400',
  fontSize: '16px',
  color: '#000',
  marginTop: '16px',
  '@bp3': {
    fontSize: '14px',
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
  '@bp3': {
    fontSize: '14px',
  },
})
