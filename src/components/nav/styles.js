import { styled } from 'config/stitches.config'

export const Nav = styled('div', {
  width: '100%',
  height: '80px',
  background: '$neutral1',
  padding: '0 16px',
  ul: {
    padding: '0',
    li: {
      display: 'inline-block',
      a: {
        color: '$neutral1',
      },
    },
  },
  '@bp2': {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    height: '58px',
    boxShadow: '0px 1px 2px rgba(0, 0, 0, 0.25)',
  },
})

export const Brand = styled('div', {
  width: '100px',
  paddingTop: '12px',
  margin: 'auto',
  img: {
    width: '100%',
  },
  '@bp2': {
    width: '60px',
    paddingTop: '2px',
  },
})

export const Menu = styled('div', {
  width: '100%',
  height: '100%',
  display: 'none',
  background: '#fff',
  position: 'absolute',
  zIndex: '1000',
  padding: '42px 32px',
  transition: 'transform .3s ease',
  overflow: 'hidden',
  color: '#000',
  ul: {
    margin: '0',
    padding: '0',
    li: {
      display: 'block',
      a: {
        padding: '32px 16px',
        display: 'block',
        cursor: 'pointer',
        borderRadius: '4px',
        '&:hover': {
          color: '#fff',
          background: '$primary1',
        },
      },
    },
  },
  '@bp2': {
    display: 'block',
  },
})

export const Close = styled('div', {
  width: '28px',
  height: '32px',
  position: 'absolute',
  top: '16px',
  right: '16px',
  cursor: 'pointer',
})

export const Avatar = styled('div', {
  width: '32px',
  height: '32px',
  display: 'none',
  '@bp3': {
    display: 'flex',
    alignItems: 'center',
  },
})

export const User = styled('div', {
  width: '300px',
  height: '80px',
  paddingRight: '16px',
  position: 'absolute',
  top: '0',
  right: '0',
  textAlign: 'right',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'flex-end',
  fontSize: '14px',
  cursor: 'pointer',
  span: {
    fontWeight: '600',
  },
  '@bp2': {
    display: 'none',
  },
})

export const Button = styled('button', {
  width: '120px',
  borderRadius: 4,
  padding: '8px 16px',
  background: '#fff',
  color: '$primary',
  fontSize: 18,
  lineHeight: 1,
  fontWeight: 500,
  cursor: 'pointer',
  border: '1x solid #840B09',
})

export const Mobile = styled('div', {
  width: '32px',
  height: '32px',
  display: 'none',
  margin: '0 0px 0 0px',
  textAlign: 'center',
  paddingTop: '2px',
  cursor: 'pointer',
  '@bp2': {
    display: 'block',
  },
})
