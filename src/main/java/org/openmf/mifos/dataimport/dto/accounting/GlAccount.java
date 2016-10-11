package org.openmf.mifos.dataimport.dto.accounting;

import org.openmf.mifos.dataimport.dto.Type;


public class GlAccount {
	private final Integer id;

	private final String name;

	private final Type usage;

	public GlAccount(Integer id, String name, Type usage) {

		this.id = id;
		this.name = name;
		this.usage = usage;

	}

	public Integer getId() {
		return id;
	}

	public String getName() {
		return name;
	}

	public Type getUsage() {
		return usage;
	}
}

