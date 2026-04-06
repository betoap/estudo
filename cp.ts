offers: OfferItem[] = [
  {
    id: 1,
    type: 'negotiated',
    badge: 'Taxa negociada',
    title: 'Aplicação Financeira',
    guaranteePercentage: '100%',

    availability: 'R$ 150.000,00',
    contractTerm: '12x R$ 13.240,00',
    insurance: 'Incluído R$ 25,47 a.m. do seguro',

    firstInstallmentDate: '14 jan 2026',
    annualSpread: '2,67% a.m.',
    interestRate: '1,99% a.m. / 26,67% a.a.',

    actionLabel: 'Contratar',
    actionType: 'contract',
  },
  {
    id: 2,
    type: 'preApproved',
    badge: 'Pré-aprovado',
    title: 'Devedor solidário',
    guaranteePercentage: '100%',

    availability: 'R$ 80.000,00',
    contractTerm: 'Entre 6 a 60 meses',

    additionalLabel: 'Transferência de limite',
    additionalValue: 'Entre R$ 5.000,00 e R$ 80.000,00',

    actionLabel: 'Simular',
    actionType: 'simulate',
  },
  {
    id: 3,
    type: 'preApproved',
    badge: 'Pré-aprovado',
    title: 'Aplicação Financeira',
    guaranteePercentage: '100%',

    availability: 'R$ 220.000,00',
    contractTerm: 'Entre 6 a 60 meses',

    additionalLabel: 'Investimentos',
    additionalValue: 'Dos sócios e da empresa, exceto conta conjunta',

    actionLabel: 'Simular',
    actionType: 'simulate',
  }
];


get negotiatedOffers(): OfferItem[] {
  return this.offers.filter(item => item.type === 'negotiated');
}

get preApprovedOffers(): OfferItem[] {
  return this.offers.filter(item => item.type === 'preApproved');
}


get negotiatedOffers(): OfferItem[] {
    return this.offers.filter(item => item.type === 'negotiated');
  }

  get preApprovedOffers(): OfferItem[] {
    return this.offers.filter(item => item.type === 'preApproved');
  }

  trackOffer(index: number, offer: OfferItem): number | string {
    return offer.id;
  }
